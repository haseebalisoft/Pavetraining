import "server-only";

import {
  asNullableString,
  createListItemByKey,
  extractLookupId,
  getListItemsByKey,
  updateListItemFieldsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { CLIENT_MATRIX_CATEGORY_COLUMNS } from "@/lib/services/bulkUpload/clientTemplateHeaders";
import { normalizeDateValue } from "@/lib/services/bulkUpload/parseSpreadsheet";

const fields = getSharePointFields("trainingMatrixCategoryRecords");
const nporsCategoryFields = getSharePointFields("nporsCategories");
const workforceFields = getSharePointFields("workforce");
const companyFields = getSharePointFields("company");

export type MatrixCategoryWrite = {
  code: string;
  name: string;
  expiryDate: string;
};

export type MatrixCategoryLookupCaches = {
  workforceByName: Map<string, string>;
  companyByName: Map<string, string>;
  categoryByCode: Map<string, string>;
  /** Existing category-record keys → item id (mutated as rows are written). */
  existingByKey?: Map<string, string>;
  /** Same keys → stored expiry (ISO date), so unchanged records are not re-written. */
  existingExpiryByKey?: Map<string, string>;
};

/**
 * Category-record writes in flight per matrix row. Each spreadsheet row
 * carries up to ~150 N-code dates; one-at-a-time writes made a 50-row upload
 * take ~90 minutes. Kept modest so SharePoint throttling (429, retried by the
 * Graph SDK) stays rare.
 */
const CATEGORY_WRITE_CONCURRENCY = 6;

/** Run `work` over `items` with at most `limit` in flight. */
async function forEachLimited<T>(
  items: T[],
  limit: number,
  work: (item: T) => Promise<void>,
): Promise<void> {
  let next = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const item = items[next];
      next += 1;
      await work(item!);
    }
  });
  await Promise.all(runners);
}

function nameKey(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

/** Pull short code from NPORS Categories Title like "N001 - Industrial…". */
export function extractCategoryCode(
  value: string | null | undefined,
): string | null {
  if (!value?.trim()) return null;
  const text = value.trim();
  const meta = ["CSCS", "SSSTS", "SMSTS", "NRSWA", "EUSR", "FACEFIT"];
  const upper = text.toUpperCase();
  for (const code of meta) {
    if (
      upper === code ||
      upper.startsWith(`${code} `) ||
      upper.startsWith(`${code}-`)
    ) {
      return code;
    }
  }
  const match = text.match(/^(N\d+[A-Z]?)/i);
  return match?.[1]?.toUpperCase() ?? upper;
}

async function loadLookupCaches(): Promise<MatrixCategoryLookupCaches> {
  // These four scans return the FULL Workforce / Company / NPORS Categories
  // / Matrix Category Records lists and combined can easily exceed the
  // Next.js unstable_cache 2MB limit — which throws an unhandledRejection
  // when it tries to persist the result. Read uncached (the bulk import is
  // a one-shot job, so cache re-use gives us nothing here).
  const [workforceItems, companyItems, categoryItems, existingRecords] =
    await Promise.all([
      getListItemsByKey("workforce", { top: 5000, skipCache: true }),
      getListItemsByKey("company", { top: 5000, skipCache: true }),
      getListItemsByKey("nporsCategories", {
        top: 5000,
        skipCache: true,
      }).catch(() => []),
      // A full matrix upload creates ~150 records per candidate; a 5000 cap
      // hid existing records past it, so re-uploads created duplicates.
      getListItemsByKey("trainingMatrixCategoryRecords", {
        top: 100_000,
        skipCache: true,
      }).catch(() => []),
    ]);

  const workforceByName = new Map<string, string>();
  for (const item of workforceItems) {
    const candidateName =
      asNullableString(item.fields[workforceFields.candidateName]) ??
      asNullableString(item.fields.Title);
    if (!candidateName) continue;
    const key = nameKey(candidateName);
    if (!workforceByName.has(key)) workforceByName.set(key, item.id);
  }

  const companyByName = new Map<string, string>();
  for (const item of companyItems) {
    const companyName =
      asNullableString(item.fields[companyFields.companyName]) ??
      asNullableString(item.fields[companyFields.title]);
    if (!companyName) continue;
    companyByName.set(nameKey(companyName), item.id);
  }

  const categoryByCode = new Map<string, string>();
  for (const item of categoryItems) {
    const title =
      asNullableString(item.fields[nporsCategoryFields.title]) ??
      asNullableString(item.fields[nporsCategoryFields.categoryCode]);
    const code = extractCategoryCode(title);
    if (!code) continue;
    if (!categoryByCode.has(code)) categoryByCode.set(code, item.id);
  }

  const existingByKey = new Map<string, string>();
  const existingExpiryByKey = new Map<string, string>();
  for (const item of existingRecords) {
    const cand = extractLookupId(item.fields, fields.candidateName);
    const comp = extractLookupId(item.fields, fields.companyName);
    const cat = extractLookupId(item.fields, fields.categoryCode);
    if (cand && comp && cat) {
      const key = `${cand}|${comp}|${cat}`;
      existingByKey.set(key, item.id);
      const expiry = normalizeDateValue(
        asNullableString(item.fields[fields.expiryDate]),
      );
      if (expiry) existingExpiryByKey.set(key, expiry);
    }
  }

  return {
    workforceByName,
    companyByName,
    categoryByCode,
    existingByKey,
    existingExpiryByKey,
  };
}

async function ensureCategoryCode(
  code: string,
  name: string,
  cache: MatrixCategoryLookupCaches,
): Promise<string> {
  const key = code.trim().toUpperCase();
  const existing = cache.categoryByCode.get(key);
  if (existing) return existing;

  const title = /^N\d+/i.test(key) ? `${key} - ${name}` : key;
  const created = await createListItemByKey("nporsCategories", {
    [nporsCategoryFields.title]: title,
    [nporsCategoryFields.categoryCode]: key,
    [nporsCategoryFields.active]: true,
    [nporsCategoryFields.customerVisible]: true,
  });
  cache.categoryByCode.set(key, created.id);
  return created.id;
}

/**
 * Upsert category expiry rows for one candidate/company.
 * Candidate / Company / Category Code are Lookups — write LookupIds only.
 */
export async function upsertMatrixCategoryRecords(input: {
  candidateName: string;
  companyName: string;
  /**
   * Workforce / Company item ids the caller already matched (Name + DOB).
   * Always pass them when known: the name lookup below picks the FIRST
   * same-named person, so two people called e.g. "Aiden Mercer" at one
   * company had their dates written onto the same records.
   */
  candidateId?: string | null;
  companyId?: string | null;
  categories: MatrixCategoryWrite[];
  caches?: MatrixCategoryLookupCaches;
}): Promise<{
  written: number;
  unchanged: number;
  failed: number;
  errors: string[];
}> {
  const caches = input.caches ?? (await loadLookupCaches());
  const byKey =
    caches.existingByKey ??
    ((caches.existingByKey = new Map<string, string>()), caches.existingByKey);

  const candidateId =
    input.candidateId?.trim() ||
    caches.workforceByName.get(nameKey(input.candidateName));
  const companyId =
    input.companyId?.trim() ||
    caches.companyByName.get(nameKey(input.companyName));
  if (!candidateId) {
    return {
      written: 0,
      unchanged: 0,
      failed: input.categories.length,
      errors: [
        `Candidate "${input.candidateName}" was not found in Workforce for category writes.`,
      ],
    };
  }
  if (!companyId) {
    return {
      written: 0,
      unchanged: 0,
      failed: input.categories.length,
      errors: [
        `Company "${input.companyName}" was not found for category writes.`,
      ],
    };
  }

  const expiryByKey =
    caches.existingExpiryByKey ??
    ((caches.existingExpiryByKey = new Map<string, string>()),
    caches.existingExpiryByKey);
  const keyOf = (candId: string, compId: string, catId: string) =>
    `${candId}|${compId}|${catId}`;

  let written = 0;
  let failed = 0;
  let unchanged = 0;
  const errors: string[] = [];

  // Resolve category codes first, serially: a missing code is CREATED in the
  // NPORS Categories list, and two parallel writers must not create it twice.
  const work: Array<{ key: string; payload: SharePointFields; expiry: string }> = [];
  for (const category of input.categories) {
    const expiry = normalizeDateValue(category.expiryDate);
    if (!expiry) continue;

    let categoryId: string;
    try {
      categoryId = await ensureCategoryCode(
        category.code,
        category.name,
        caches,
      );
    } catch (error) {
      failed += 1;
      errors.push(
        `${category.code}: could not resolve NPORS category (${
          error instanceof Error ? error.message : "unknown"
        })`,
      );
      continue;
    }

    const key = keyOf(candidateId, companyId, categoryId);
    // Same date already stored: nothing to write. Re-uploading a sheet (or
    // resuming one that hit the time limit) now costs reads, not writes.
    if (byKey.has(key) && expiryByKey.get(key) === expiry) {
      unchanged += 1;
      continue;
    }

    const title = `${input.candidateName} · ${category.code}`.slice(0, 240);
    work.push({
      key,
      expiry,
      payload: {
        [fields.title]: title,
        [`${fields.candidateName}LookupId`]: Number(candidateId),
        [`${fields.companyName}LookupId`]: Number(companyId),
        [`${fields.categoryCode}LookupId`]: Number(categoryId),
        [fields.categoryName]: category.name,
        [fields.expiryDate]: expiry,
        [fields.status]: "Active",
        [fields.customerVisible]: true,
      },
    });
  }

  await forEachLimited(work, CATEGORY_WRITE_CONCURRENCY, async (item) => {
    const existingId = byKey.get(item.key);
    try {
      if (existingId) {
        // skipReload: the caller never reads the item back, and the reload
        // GET doubled the Graph calls per category.
        await updateListItemFieldsByKey(
          "trainingMatrixCategoryRecords",
          existingId,
          item.payload,
          { skipReload: true },
        );
      } else {
        const created = await createListItemByKey(
          "trainingMatrixCategoryRecords",
          item.payload,
        );
        byKey.set(item.key, created.id);
      }
      expiryByKey.set(item.key, item.expiry);
      written += 1;
    } catch (error) {
      failed += 1;
      errors.push(
        `${String(item.payload[fields.title] ?? item.key)}: ${
          error instanceof Error ? error.message : "write failed"
        }`,
      );
    }
  });

  return { written, unchanged, failed, errors };
}

/** Warm lookup caches once for a bulk matrix import. */
export async function loadMatrixCategoryLookupCaches(): Promise<MatrixCategoryLookupCaches> {
  return loadLookupCaches();
}

/** Extract N-code category dates from a client matrix spreadsheet row. */
export function extractCategoryWritesFromRow(
  raw: Record<string, string | null>,
): MatrixCategoryWrite[] {
  const writes: MatrixCategoryWrite[] = [];
  const byHeader = new Map<string, string | null>();
  for (const [key, value] of Object.entries(raw)) {
    byHeader.set(key.trim().toLowerCase().replace(/\s+/g, " "), value);
  }

  for (const column of CLIENT_MATRIX_CATEGORY_COLUMNS) {
    const value =
      byHeader.get(column.header.trim().toLowerCase().replace(/\s+/g, " ")) ??
      raw[column.header] ??
      null;
    const expiry = normalizeDateValue(value);
    if (!expiry) continue;
    writes.push({
      code: column.code,
      name: column.name,
      expiryDate: expiry,
    });
  }

  const meta: Array<{ code: string; name: string; aliases: string[] }> = [
    { code: "CSCS", name: "CSCS", aliases: ["CSCS Expiry"] },
    { code: "SSSTS", name: "SSSTS", aliases: ["SSSTS Expiry"] },
    { code: "SMSTS", name: "SMSTS", aliases: ["SMSTS Expiry"] },
    { code: "NRSWA", name: "NRSWA", aliases: ["NRSWA Expiry"] },
    { code: "EUSR", name: "EUSR", aliases: ["EUSR Expiry"] },
    {
      code: "FACEFIT",
      name: "Face Fit",
      aliases: ["Face ift", "Face Fit", "FaceFit"],
    },
  ];
  for (const item of meta) {
    let value: string | null = null;
    for (const alias of item.aliases) {
      value =
        byHeader.get(alias.trim().toLowerCase().replace(/\s+/g, " ")) ?? null;
      if (value) break;
    }
    const expiry = normalizeDateValue(value);
    if (!expiry) continue;
    writes.push({
      code: item.code,
      name: item.name,
      expiryDate: expiry,
    });
  }

  return writes;
}

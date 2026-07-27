import "server-only";

import {
  createListItemByKey,
  getListItemsByKey,
  updateListItemFieldsByKey,
} from "@/lib/services/sharePointListService";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { CLIENT_MATRIX_CATEGORY_COLUMNS } from "@/lib/services/bulkUpload/clientTemplateHeaders";
import { normalizeDateValue } from "@/lib/services/bulkUpload/parseSpreadsheet";

const fields = getSharePointFields("trainingMatrixCategoryRecords");

export type MatrixCategoryWrite = {
  code: string;
  name: string;
  expiryDate: string;
};

/**
 * Upsert category expiry rows for one candidate/company.
 * Matches existing rows by candidate + company + category code.
 */
export async function upsertMatrixCategoryRecords(input: {
  candidateName: string;
  companyName: string;
  categories: MatrixCategoryWrite[];
}): Promise<{ written: number; failed: number; errors: string[] }> {
  const existing = await getListItemsByKey("trainingMatrixCategoryRecords", {
    top: 5000,
  }).catch(() => []);

  const keyOf = (candidate: string, company: string, code: string) =>
    `${candidate.trim().toLowerCase()}|${company.trim().toLowerCase()}|${code.trim().toLowerCase()}`;

  const byKey = new Map<string, string>();
  for (const item of existing) {
    const candidate = String(item.fields[fields.candidateName] ?? "").trim();
    const company = String(item.fields[fields.companyName] ?? "").trim();
    const code = String(item.fields[fields.categoryCode] ?? "").trim();
    if (candidate && company && code) {
      byKey.set(keyOf(candidate, company, code), item.id);
    }
  }

  let written = 0;
  let failed = 0;
  const errors: string[] = [];

  for (const category of input.categories) {
    const expiry = normalizeDateValue(category.expiryDate);
    if (!expiry) continue;
    const title = `${input.candidateName} · ${category.code}`.slice(0, 240);
    const payload = {
      [fields.title]: title,
      [fields.candidateName]: input.candidateName,
      [fields.companyName]: input.companyName,
      [fields.categoryCode]: category.code,
      [fields.categoryName]: category.name,
      [fields.expiryDate]: expiry,
      [fields.status]: "Active",
      [fields.customerVisible]: true,
    };
    const existingId = byKey.get(
      keyOf(input.candidateName, input.companyName, category.code),
    );
    try {
      if (existingId) {
        await updateListItemFieldsByKey(
          "trainingMatrixCategoryRecords",
          existingId,
          payload,
        );
      } else {
        const created = await createListItemByKey(
          "trainingMatrixCategoryRecords",
          payload,
        );
        byKey.set(
          keyOf(input.candidateName, input.companyName, category.code),
          created.id,
        );
      }
      written += 1;
    } catch (error) {
      failed += 1;
      errors.push(
        `${category.code}: ${
          error instanceof Error ? error.message : "write failed"
        }`,
      );
    }
  }

  return { written, failed, errors };
}

/** Extract N-code category dates from a client matrix spreadsheet row. */
export function extractCategoryWritesFromRow(
  raw: Record<string, string | null>,
): MatrixCategoryWrite[] {
  const writes: MatrixCategoryWrite[] = [];
  const byHeader = new Map<string, string | null>();
  for (const [key, value] of Object.entries(raw)) {
    byHeader.set(key.trim().toLowerCase(), value);
  }

  for (const column of CLIENT_MATRIX_CATEGORY_COLUMNS) {
    const value =
      byHeader.get(column.header.trim().toLowerCase()) ??
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

  // Meta card-style columns also stored as category records.
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
      value = byHeader.get(alias.trim().toLowerCase()) ?? null;
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

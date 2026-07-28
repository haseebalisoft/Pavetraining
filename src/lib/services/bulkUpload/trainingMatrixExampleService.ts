import "server-only";

import {
  cachedSharePointRead,
  sharePointListTag,
} from "@/lib/cache/sharePointCache";
import { getSharePointListId, getSharePointSiteApiRoot } from "@/lib/config/sharepoint";
import { getGraphClient } from "@/lib/graph/graphClient";
import { SHAREPOINT_LISTS } from "@/lib/schema/sharepointSchema";
import { CLIENT_MATRIX_DISPLAY_HEADERS } from "@/lib/services/bulkUpload/clientTemplateHeaders";
import {
  getListItemsByKey,
  createListItemByKey,
  updateListItemFieldsByKey,
  type SharePointListItem,
} from "@/lib/services/sharePointListService";
import { normalizeDateValue } from "@/lib/services/bulkUpload/parseSpreadsheet";

export interface TrainingMatrixExampleRow {
  id: string;
  candidateName: string;
  dateOfBirth: string | null;
  columnValues: Record<string, string | null>;
  nextExpiryDate: string | null;
}

const EXAMPLE_LIST_ENV =
  SHAREPOINT_LISTS.trainingMatrixExample.listIdEnvVar;
const EXAMPLE_LIST_DISPLAY_NAME =
  SHAREPOINT_LISTS.trainingMatrixExample.displayName;

/**
 * Resolve Training Matrix Update list GUID from env, or look it up on SharePoint
 * by display name and cache onto process.env for later getSharePointListId calls.
 */
export async function resolveTrainingMatrixExampleListId(): Promise<
  string | null
> {
  const fromEnv = process.env[EXAMPLE_LIST_ENV]?.trim();
  if (fromEnv) return fromEnv;

  try {
    const siteRoot = getSharePointSiteApiRoot();
    const client = getGraphClient();
    const escaped = EXAMPLE_LIST_DISPLAY_NAME.replace(/'/g, "''");
    const response = (await client
      .api(`${siteRoot}/lists`)
      .filter(`displayName eq '${escaped}'`)
      .select("id,displayName")
      .top(5)
      .get()) as { value?: Array<{ id?: string; displayName?: string }> };

    const match = response.value?.find(
      (list) =>
        list.displayName?.trim().toLowerCase() ===
        EXAMPLE_LIST_DISPLAY_NAME.trim().toLowerCase(),
    );
    const id = match?.id?.trim() || null;
    if (id) {
      process.env[EXAMPLE_LIST_ENV] = id;
    }
    return id;
  } catch {
    return null;
  }
}

function normalizeHeader(value: string): string {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

/** Excel serial day number (SharePoint "From Excel" import) → YYYY-MM-DD. */
export function excelSerialToIsoDate(value: unknown): string | null {
  if (value == null || value === "") return null;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed || /^(—|–|-|n\/?a|null|none|0)$/i.test(trimmed)) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) return trimmed.slice(0, 10);
    const asNum = Number(trimmed);
    if (!Number.isNaN(asNum) && asNum > 20000) {
      return excelSerialToIsoDate(asNum);
    }
    const parsed = Date.parse(trimmed);
    if (!Number.isNaN(parsed)) {
      return new Date(parsed).toISOString().slice(0, 10);
    }
    return null;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value) || value === 0) return null;
    // Excel epoch 1899-12-30 (Graph/SharePoint Excel import serials)
    const ms = Date.UTC(1899, 11, 30) + Math.round(value) * 86_400_000;
    return new Date(ms).toISOString().slice(0, 10);
  }
  return null;
}

function earliestDate(values: Array<string | null | undefined>): string | null {
  let min: number | null = null;
  let iso: string | null = null;
  for (const value of values) {
    if (!value?.trim()) continue;
    const t = new Date(value).getTime();
    if (Number.isNaN(t)) continue;
    if (min == null || t < min) {
      min = t;
      iso = value.slice(0, 10);
    }
  }
  return iso;
}

export function earliestDateFromColumns(
  columnValues: Record<string, string | null>,
): string | null {
  return earliestDate(
    Object.entries(columnValues)
      .filter(([key]) => key !== "Name" && key !== "DOB")
      .map(([, value]) => value),
  );
}

type ExampleColumnInfo = {
  name: string;
  /** SharePoint Number columns store Excel serials; DateTime columns store ISO. */
  storage: "number" | "dateTime" | "other";
};

/** Collapse dashes/degrees so Excel headers match SharePoint display names. */
function headerLookupKey(value: string): string {
  return normalizeHeader(value)
    .toLowerCase()
    .replace(/[–—]/g, "-")
    .replace(/[°º]/g, "");
}

/**
 * Column map must be a plain object for Next.js unstable_cache (JSON).
 * Returning a Map serializes to {} and later .get() throws → empty matrix UI.
 */
async function getExampleColumnRecords(): Promise<
  Record<string, ExampleColumnInfo>
> {
  const listId =
    process.env[EXAMPLE_LIST_ENV]?.trim() ??
    (await resolveTrainingMatrixExampleListId());
  if (!listId) {
    return {};
  }
  const siteRoot = getSharePointSiteApiRoot();

  return cachedSharePointRead(
    ["sp-matrix-example-columns-v3", listId],
    [sharePointListTag("trainingMatrixExample")],
    async () => {
      const client = getGraphClient();
      const records: Record<string, ExampleColumnInfo> = {};
      let url = `${siteRoot}/lists/${listId}/columns?$top=200`;
      while (url) {
        const page = (await client.api(url).get()) as {
          value?: Array<{
            name?: string;
            displayName?: string;
            readOnly?: boolean;
            number?: unknown;
            dateTime?: unknown;
          }>;
          "@odata.nextLink"?: string;
        };
        for (const col of page.value ?? []) {
          if (col.readOnly || !col.name || !col.displayName) continue;
          if (col.name === "ContentType" || col.name === "Attachments") continue;
          const info: ExampleColumnInfo = {
            name: col.name,
            storage: col.dateTime
              ? "dateTime"
              : col.number
                ? "number"
                : "other",
          };
          const key = headerLookupKey(col.displayName);
          records[key] = info;
          // Also index by exact lower display for simple headers like "dob"
          records[normalizeHeader(col.displayName).toLowerCase()] = info;
          if (col.name === "Title") {
            records.name = info;
          }
        }
        url = page["@odata.nextLink"]
          ? page["@odata.nextLink"].replace(
              /^https:\/\/graph\.microsoft\.com\/v1\.0/i,
              "",
            )
          : "";
      }
      return records;
    },
    300,
  );
}

async function getExampleColumnMap(): Promise<Map<string, ExampleColumnInfo>> {
  const records = await getExampleColumnRecords();
  return new Map(Object.entries(records));
}

function mapItemToRow(
  item: SharePointListItem,
  columnMap: Map<string, ExampleColumnInfo>,
): TrainingMatrixExampleRow | null {
  const fields = item.fields ?? {};
  const nameInternal = columnMap.get("name")?.name;
  const title =
    typeof fields.Title === "string"
      ? fields.Title.trim()
      : typeof fields.LinkTitle === "string"
        ? fields.LinkTitle.trim()
        : typeof fields.CandidateNameText === "string"
          ? fields.CandidateNameText.trim()
          : nameInternal && typeof fields[nameInternal] === "string"
            ? String(fields[nameInternal]).trim()
            : "";
  if (!title) return null;

  const columnValues: Record<string, string | null> = { Name: title };

  for (const header of CLIENT_MATRIX_DISPLAY_HEADERS) {
    if (header === "Name") continue;
    const col =
      columnMap.get(headerLookupKey(header)) ??
      columnMap.get(normalizeHeader(header).toLowerCase()) ??
      null;
    if (!col) {
      columnValues[header] = null;
      continue;
    }
    const raw = fields[col.name];
    columnValues[header] = excelSerialToIsoDate(raw);
  }

  const dateOfBirth = columnValues.DOB ?? null;
  const nextExpiryDate = earliestDate(
    CLIENT_MATRIX_DISPLAY_HEADERS.filter((h) => h !== "Name" && h !== "DOB").map(
      (h) => columnValues[h],
    ),
  );

  return {
    id: item.id,
    candidateName: title,
    dateOfBirth,
    columnValues,
    nextExpiryDate,
  };
}

/**
 * Reads SharePoint "Training Matrix Update" (wide DateTime matrix list).
 * Returns [] only when the list id cannot be resolved.
 */
export async function listTrainingMatrixExampleRows(): Promise<
  TrainingMatrixExampleRow[]
> {
  const listId = await resolveTrainingMatrixExampleListId();
  if (!listId) {
    return [];
  }

  try {
    const items = await getListItemsByKey("trainingMatrixExample", {
      top: 5000,
    });
    const columnMap = await getExampleColumnMap().catch((error) => {
      console.error(
        "[trainingMatrixExample] column map failed; showing names only:",
        error,
      );
      return new Map<string, ExampleColumnInfo>();
    });
    return items
      .map((item) => mapItemToRow(item, columnMap))
      .filter((row): row is TrainingMatrixExampleRow => Boolean(row));
  } catch (error) {
    console.error("[trainingMatrixExample] list rows failed:", error);
    throw error;
  }
}

export function mergeExampleColumnValues(
  base: Record<string, string | null>,
  example: Record<string, string | null> | undefined,
): Record<string, string | null> {
  if (!example) return base;
  const merged = { ...base };
  for (const [key, value] of Object.entries(example)) {
    if (key === "Name") continue;
    if (value?.trim()) merged[key] = value;
  }
  return merged;
}

/** ISO date → SharePoint DateTime payload (`YYYY-MM-DDT00:00:00Z`). */
export function isoToSharePointDateTime(
  iso: string | null | undefined,
): string | null {
  if (!iso?.trim()) return null;
  const text = iso.trim().slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  return `${text}T00:00:00Z`;
}

/** ISO date → Excel serial day number (for Number columns from Excel import). */
export function isoToExcelSerial(iso: string | null | undefined): number | null {
  if (!iso?.trim()) return null;
  const text = iso.trim().slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  const [y, m, d] = text.split("-").map(Number);
  const ms = Date.UTC(y, m - 1, d);
  if (Number.isNaN(ms)) return null;
  return Math.round((ms - Date.UTC(1899, 11, 30)) / 86_400_000);
}

/**
 * Create or update a row on the Excel-imported Training matrix example list.
 * Number columns (Excel import) get serials; DateTime columns get ISO dates.
 */
export async function upsertTrainingMatrixExampleRow(input: {
  candidateName: string;
  /** Existing SharePoint item id (without `example:` prefix), if updating. */
  existingItemId?: string | null;
  /** Spreadsheet cells keyed by template headers (Name, DOB, CSCS Expiry, N001 - …). */
  source: Record<string, string | null>;
}): Promise<{ id: string; created: boolean }> {
  const name = input.candidateName.trim();
  if (!name) throw new Error("Candidate name is required for matrix example upsert.");

  const listId = await resolveTrainingMatrixExampleListId();
  if (!listId) {
    throw new Error(
      `SharePoint list "${EXAMPLE_LIST_DISPLAY_NAME}" not found. Set ${EXAMPLE_LIST_ENV} or create the list.`,
    );
  }

  const columnMap = await getExampleColumnMap();
  const fields: Record<string, unknown> = { Title: name };

  for (const header of CLIENT_MATRIX_DISPLAY_HEADERS) {
    if (header === "Name") continue;
    const col = columnMap.get(headerLookupKey(header));
    if (!col) continue;

    const raw =
      input.source[header] ??
      input.source[normalizeHeader(header)] ??
      null;
    const asText =
      raw == null ? null : typeof raw === "string" ? raw : String(raw);
    const iso =
      normalizeDateValue(asText) ?? excelSerialToIsoDate(raw);

    if (col.storage === "dateTime") {
      fields[col.name] = isoToSharePointDateTime(iso);
    } else {
      // Excel-imported Number columns must receive serial day numbers.
      fields[col.name] = isoToExcelSerial(iso) ?? 0;
    }
  }

  if (input.existingItemId?.trim()) {
    await updateListItemFieldsByKey(
      "trainingMatrixExample",
      input.existingItemId.trim(),
      fields,
    );
    return { id: input.existingItemId.trim(), created: false };
  }

  const created = await createListItemByKey("trainingMatrixExample", fields);
  return { id: created.id, created: true };
}

export function stripExampleMatrixId(id: string | null | undefined): string | null {
  if (!id?.trim()) return null;
  return id.startsWith("example:") ? id.slice("example:".length) : id.trim();
}

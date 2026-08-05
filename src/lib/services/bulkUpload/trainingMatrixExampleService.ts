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
import {
  buildWorkforceMatrixSource,
  findMatrixRowForCandidate,
  realMatrixItemId,
  type WorkforceMatrixProfile,
} from "@/lib/services/bulkUpload/workforceMatrixSync";
import {
  MANUAL_OVERRIDES_FIELD,
  parseManualOverrides,
  serializeManualOverrides,
} from "@/lib/training/matrixManualOverrides";

export interface TrainingMatrixExampleRow {
  id: string;
  candidateName: string;
  dateOfBirth: string | null;
  columnValues: Record<string, string | null>;
  nextExpiryDate: string | null;
  /** Headers whose dates were set manually in admin (not register sync). */
  manualOverrides: string[];
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

  const directFieldByHeader: Record<string, string> = {
    "CSCS Expiry": "CSCSExpiry",
    "SSSTS Expiry": "SSSTSExpiry",
    "SMSTS Expiry": "SMSTSExpiry",
    "NRSWA Expiry": "NRSWAExpiry",
    "EUSR Expiry": "EUSRExpiry",
    DOB: "DOB",
    "N001 - Ind FLT": "N001_x002d_IndFLT",
    "N003 - Reach Lift Truck": "N003_x002d_ReachLiftTruck",
    "N004 - Lorry Mounted Lift Truck": "N004_x002d_LorryMountedLiftTruck",
    "N010 - Telescopic Handler": "N010_x002d_TelescopicHandler",
    "N020 - Tiltrotator System": "N020_x002d_TiltrotatorSystem",
    "N021 - Suction Excavator": "N021_x002d_SuctionExcavator",
    "N027 - Excavation Marshal - Banksperson":
      "N027_x002d_ExcavationMarshalBanksperson",
    "N100 - Exc Crane": "N100_x002d_ExcCrane",
  };

  for (const header of CLIENT_MATRIX_DISPLAY_HEADERS) {
    if (header === "Name") continue;
    const col =
      columnMap.get(headerLookupKey(header)) ??
      columnMap.get(normalizeHeader(header).toLowerCase()) ??
      null;
    const directName = directFieldByHeader[header];
    const raw = col
      ? fields[col.name]
      : directName
        ? fields[directName]
        : undefined;
    if (raw === undefined && !col) {
      columnValues[header] = null;
      continue;
    }
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
    manualOverrides: parseManualOverrides(fields[MANUAL_OVERRIDES_FIELD]),
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
  /**
   * Non-date profile columns keyed by display name (Company, Workforce Number,
   * Training Manager, …). Each is written only when the list actually has that
   * (text) column and the value is non-blank — so it never wipes existing data.
   */
  profileFields?: Record<string, string | null>;
  /** When set, replaces ManualOverrides on the row. */
  manualOverrides?: string[] | null;
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

    // Partial updates: skip columns not present in `source` so we don't wipe dates.
    const hasKey =
      Object.prototype.hasOwnProperty.call(input.source, header) ||
      Object.prototype.hasOwnProperty.call(
        input.source,
        normalizeHeader(header),
      );
    if (!hasKey) continue;

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

  // Profile text columns (Company, Workforce Number, Training Manager, …):
  // write each only when the list has that text column and the value is
  // non-blank, so profile details appear "if fields exist" without ever
  // clobbering expiry/date columns or overwriting with blanks.
  if (input.profileFields) {
    for (const [header, value] of Object.entries(input.profileFields)) {
      if (value == null || String(value).trim() === "") continue;
      const col =
        columnMap.get(headerLookupKey(header)) ??
        columnMap.get(normalizeHeader(header).toLowerCase());
      if (!col || col.name === "Title" || col.storage !== "other") continue;
      fields[col.name] = String(value).trim();
    }
  }

  if (input.manualOverrides !== undefined) {
    fields[MANUAL_OVERRIDES_FIELD] = serializeManualOverrides(
      input.manualOverrides ?? [],
    );
  }

  if (input.existingItemId?.trim()) {
    try {
      await updateListItemFieldsByKey(
        "trainingMatrixExample",
        input.existingItemId.trim(),
        fields,
      );
    } catch (error) {
      // Retry without ManualOverrides if the column is not provisioned yet.
      if (
        input.manualOverrides !== undefined &&
        MANUAL_OVERRIDES_FIELD in fields
      ) {
        const { [MANUAL_OVERRIDES_FIELD]: _drop, ...lean } = fields;
        await updateListItemFieldsByKey(
          "trainingMatrixExample",
          input.existingItemId.trim(),
          lean,
        );
        console.warn(
          `[matrix] ManualOverrides column missing on Training Matrix Update — date saved without override flag.`,
        );
      } else {
        throw error;
      }
    }
    return { id: input.existingItemId.trim(), created: false };
  }

  try {
    const created = await createListItemByKey("trainingMatrixExample", fields);
    return { id: created.id, created: true };
  } catch (error) {
    if (
      input.manualOverrides !== undefined &&
      MANUAL_OVERRIDES_FIELD in fields
    ) {
      const { [MANUAL_OVERRIDES_FIELD]: _drop, ...lean } = fields;
      const created = await createListItemByKey("trainingMatrixExample", lean);
      return { id: created.id, created: true };
    }
    throw error;
  }
}

/**
 * Resolve a matrix id to a REAL SharePoint item id for updates. Returns null
 * for blank ids AND for synthetic `workforce-only:<id>` ids so those are never
 * used as PATCH targets (callers then create a real row instead).
 */
export function stripExampleMatrixId(id: string | null | undefined): string | null {
  return realMatrixItemId(id);
}

export interface WorkforceMatrixSyncResult {
  /** Real SharePoint item id of the Training Matrix Update row. */
  id: string;
  /** Id shape used by listAdminMatrix / the admin UI (`example:<id>`). */
  matrixId: string;
  created: boolean;
  /** Lightweight row for callers to keep an in-memory cache consistent. */
  row: TrainingMatrixExampleRow;
}

/**
 * Create or update the REAL Training Matrix Update row for a workforce record.
 *
 * The single sync path shared by manual workforce create/update AND bulk
 * workforce import, so imported candidates always get a real matrix row (never
 * a synthetic `workforce-only:` placeholder) that later matrix spreadsheet
 * imports can update by real id.
 *
 * - Reuses an existing row (matched by candidate name + DOB) — no duplicates.
 * - Writes profile detail columns that exist on the list.
 * - Never overwrites existing expiry columns with blanks (blank values are
 *   omitted from the payload entirely).
 *
 * Pass `existingRows` from a single upfront read to avoid a per-row Graph fetch
 * during bulk import.
 */
export async function syncWorkforceToTrainingMatrix(
  profile: WorkforceMatrixProfile,
  options: { existingRows?: TrainingMatrixExampleRow[] } = {},
): Promise<WorkforceMatrixSyncResult> {
  const name = profile.candidateName?.trim();
  if (!name) {
    throw new Error("Candidate name is required to sync the Training Matrix row.");
  }

  const existingRows =
    options.existingRows ?? (await listTrainingMatrixExampleRows());
  const existing = findMatrixRowForCandidate(existingRows, profile);
  const { source, profileFields } = buildWorkforceMatrixSource(profile);

  const result = await upsertTrainingMatrixExampleRow({
    candidateName: name,
    existingItemId: existing?.id ?? null,
    source,
    profileFields,
  });

  const row: TrainingMatrixExampleRow = existing
    ? existing
    : {
        id: result.id,
        candidateName: name,
        dateOfBirth: source.DOB ?? null,
        columnValues: { ...source },
        nextExpiryDate: earliestDateFromColumns(source),
        manualOverrides: [],
      };

  return {
    id: result.id,
    matrixId: `example:${result.id}`,
    created: result.created,
    row,
  };
}

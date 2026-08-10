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
  AMBIGUOUS_MATRIX_MATCH_WARNING,
  buildUnlinkedMatrixSource,
  buildWorkforceMatrixSource,
  findMatrixRowByWorkforce,
  mergeUploadedCells,
  realMatrixItemId,
  type MatrixLinkFields,
  type MatrixMatchType,
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
  /** Strong link columns (present once the row is linked to a Workforce record). */
  workforceItemId: string | null;
  workforceNumber: string | null;
  companyItemId: string | null;
  companyNumber: string | null;
  matrixLinkStatus: string | null;
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

  // Strong link columns, read by display name against the live column map so a
  // list without them (older sites) simply reports null → treated as unlinked.
  const readLinkField = (header: string): string | null => {
    const col =
      columnMap.get(headerLookupKey(header)) ??
      columnMap.get(normalizeHeader(header).toLowerCase());
    if (!col) return null;
    const raw = fields[col.name];
    if (raw == null || raw === "") return null;
    const text = String(raw).trim();
    return text ? text : null;
  };

  return {
    id: item.id,
    candidateName: title,
    dateOfBirth,
    columnValues,
    nextExpiryDate,
    manualOverrides: parseManualOverrides(fields[MANUAL_OVERRIDES_FIELD]),
    workforceItemId: readLinkField("WorkforceItemId"),
    workforceNumber: readLinkField("WorkforceNumber"),
    companyItemId: readLinkField("CompanyItemId"),
    companyNumber: readLinkField("CompanyNumber"),
    matrixLinkStatus: readLinkField("MatrixLinkStatus"),
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
  /**
   * Strong link columns keyed by display name, split by storage type:
   * `numbers` (WorkforceItemId, CompanyItemId) target Number columns, `text`
   * (WorkforceNumber, CompanyNumber, CandidateName, MatrixLinkStatus) target
   * text/choice columns. Each is written only when the column exists.
   */
  linkFields?: MatrixLinkFields;
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
      // Excel-imported Number columns store serial day numbers; a null iso is
      // an explicit clear (delete-recompute) and must stay null, not default
      // to 0 — 0 would render as a bogus 1899-12-30 "expiry".
      fields[col.name] = isoToExcelSerial(iso);
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

  // Strong link columns. Numbers → Number columns (as numeric ids); text/choice
  // → text columns. Each is written only when the list actually has that column,
  // so a site missing the link columns silently skips them (nothing breaks).
  if (input.linkFields) {
    const findCol = (header: string) =>
      columnMap.get(headerLookupKey(header)) ??
      columnMap.get(normalizeHeader(header).toLowerCase());
    for (const [header, value] of Object.entries(input.linkFields.numbers)) {
      const col = findCol(header);
      if (!col || col.name === "Title") continue;
      // Never write an id into a date column; a text link column takes a string.
      if (col.storage === "number") fields[col.name] = value;
      else if (col.storage === "other") fields[col.name] = String(value);
    }
    for (const [header, value] of Object.entries(input.linkFields.text)) {
      const text = String(value).trim();
      if (!text) continue;
      const col = findCol(header);
      if (!col || col.name === "Title" || col.storage !== "other") continue;
      fields[col.name] = text;
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

export interface WorkforceMatrixSyncOptions {
  existingRows?: TrainingMatrixExampleRow[];
  /**
   * Uploaded matrix cells keyed by canonical display header. Blanks are stripped
   * before merging, so an empty spreadsheet cell can never erase a live expiry.
   */
  uploadedCells?: Record<string, string | null>;
  /** Live Workforce records sharing this candidate's name / name+DOB. */
  workforceNamePeers?: number;
  workforceNameDobPeers?: number;
}

export interface WorkforceMatrixSyncResult {
  /** Real SharePoint item id of the row, or null when the sync was skipped. */
  id: string | null;
  /** Id shape used by listAdminMatrix / the admin UI (`example:<id>`). */
  matrixId: string | null;
  created: boolean;
  /** Lightweight row for callers to keep an in-memory cache consistent. */
  row: TrainingMatrixExampleRow | null;
  /** Which link key resolved the row. */
  matchType: MatrixMatchType;
  /** True when several unlinked rows tied — nothing was written. */
  ambiguous: boolean;
  /** True when no Graph write happened. `id`/`row` are null. */
  skipped: boolean;
  warnings: string[];
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
  options: WorkforceMatrixSyncOptions = {},
): Promise<WorkforceMatrixSyncResult> {
  const name = profile.candidateName?.trim();
  if (!name) {
    throw new Error("Candidate name is required to sync the Training Matrix row.");
  }

  const existingRows =
    options.existingRows ?? (await listTrainingMatrixExampleRows());
  // Match by the STRONG link keys (WorkforceItemId → WorkforceNumber+CompanyItemId
  // → Company+Name+DOB → Name+DOB). Rows owned by another workforce record are
  // invisible to the name/DOB steps, so a row is never hijacked.
  const match = findMatrixRowByWorkforce(existingRows, profile, {
    workforceNamePeers: options.workforceNamePeers,
    workforceNameDobPeers: options.workforceNameDobPeers,
  });
  const warnings: string[] = [];

  // Several unlinked rows tied: writing either one could attach this candidate's
  // training to the wrong person, and creating a third would duplicate. Leave
  // them Needs Review for an admin.
  if (match.ambiguous) {
    return {
      id: null,
      matrixId: null,
      created: false,
      row: null,
      matchType: "none",
      ambiguous: true,
      skipped: true,
      warnings: [AMBIGUOUS_MATRIX_MATCH_WARNING],
    };
  }

  let existing = match.row;
  // Guard a stale existingRows cache (parallel bulk import): if the matched row
  // has since been claimed by someone else, create this candidate's own row.
  if (existing) {
    const owner = String(existing.workforceItemId ?? "").trim();
    const self = String(profile.workforceItemId ?? profile.id ?? "").trim();
    if (owner && self && owner !== self) {
      warnings.push(
        `Matrix row #${existing.id} is linked to another Workforce record — a new row was created instead.`,
      );
      existing = null;
    }
  }

  const built = buildWorkforceMatrixSource(profile);
  const source = mergeUploadedCells(built.source, options.uploadedCells);
  const { profileFields, linkFields } = built;

  const result = await upsertTrainingMatrixExampleRow({
    candidateName: name,
    existingItemId: existing?.id ?? null,
    source,
    profileFields,
    linkFields,
  });

  // Overlay the just-written link values so an in-memory cache (bulk import)
  // reflects the row as linked, even when reusing an existing row object.
  const linkFieldValues = {
    workforceItemId:
      linkFields.numbers.WorkforceItemId != null
        ? String(linkFields.numbers.WorkforceItemId)
        : (existing?.workforceItemId ?? null),
    workforceNumber: linkFields.text.WorkforceNumber ?? existing?.workforceNumber ?? null,
    companyItemId:
      linkFields.numbers.CompanyItemId != null
        ? String(linkFields.numbers.CompanyItemId)
        : (existing?.companyItemId ?? null),
    companyNumber: linkFields.text.CompanyNumber ?? existing?.companyNumber ?? null,
    matrixLinkStatus: linkFields.text.MatrixLinkStatus ?? "Linked",
  };

  // Mirror what the upsert wrote so callers reusing this row as a cache entry
  // (bulk import, matrix importer) see the post-write state without re-reading.
  const columnValues = existing
    ? { ...existing.columnValues, ...source }
    : { ...source };

  const row: TrainingMatrixExampleRow = existing
    ? {
        ...existing,
        candidateName: name,
        dateOfBirth: source.DOB ?? existing.dateOfBirth ?? null,
        columnValues,
        nextExpiryDate: earliestDateFromColumns(columnValues),
        ...linkFieldValues,
      }
    : {
        id: result.id,
        candidateName: name,
        dateOfBirth: source.DOB ?? null,
        columnValues,
        nextExpiryDate: earliestDateFromColumns(columnValues),
        manualOverrides: [],
        ...linkFieldValues,
      };

  return {
    id: result.id,
    matrixId: `example:${result.id}`,
    created: result.created,
    row,
    matchType: existing ? match.matchType : "none",
    ambiguous: false,
    skipped: false,
    warnings,
  };
}

/**
 * Write a matrix row that has NO confirmed Workforce owner (Task C): the
 * spreadsheet gave us a candidate we cannot resolve, so the training data is
 * preserved but the row is flagged `Needs Review` and claims no WorkforceItemId.
 * A later Workforce create/import can adopt it via `syncWorkforceToTrainingMatrix`.
 */
export async function upsertUnlinkedMatrixRow(input: {
  candidateName: string;
  dateOfBirth?: string | null;
  companyName?: string | null;
  companyNumber?: string | null;
  uploadedCells?: Record<string, string | null>;
  existingRow?: TrainingMatrixExampleRow | null;
}): Promise<{ id: string; created: boolean; row: TrainingMatrixExampleRow }> {
  const { source, profileFields, linkFields } = buildUnlinkedMatrixSource({
    candidateName: input.candidateName,
    dateOfBirth: input.dateOfBirth,
    companyName: input.companyName,
    companyNumber: input.companyNumber,
    uploadedCells: input.uploadedCells,
  });

  const existing = input.existingRow ?? null;
  const result = await upsertTrainingMatrixExampleRow({
    candidateName: input.candidateName.trim(),
    existingItemId: existing?.id ?? null,
    source,
    profileFields,
    linkFields,
  });

  const columnValues = existing
    ? { ...existing.columnValues, ...source }
    : { ...source };

  return {
    id: result.id,
    created: result.created,
    row: {
      id: result.id,
      candidateName: input.candidateName.trim(),
      dateOfBirth: source.DOB ?? existing?.dateOfBirth ?? null,
      columnValues,
      nextExpiryDate: earliestDateFromColumns(columnValues),
      manualOverrides: existing?.manualOverrides ?? [],
      // Deliberately keeps workforceItemId null so a future Workforce sync can
      // still claim this row via the Company+Name+DOB step.
      workforceItemId: null,
      workforceNumber: null,
      companyItemId: existing?.companyItemId ?? null,
      companyNumber: linkFields.text.CompanyNumber ?? existing?.companyNumber ?? null,
      matrixLinkStatus: linkFields.text.MatrixLinkStatus ?? "Needs Review",
    },
  };
}

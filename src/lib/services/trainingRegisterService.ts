import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asLookupOrString,
  asNullableString,
  asString,
  getListItemsByKey,
  type SharePointFields,
  type SharePointListItem,
} from "@/lib/services/sharePointListService";
import type { AdminRegisterKey } from "@/lib/services/adminCrudService";

export type RegisterSource =
  | "NPORS"
  | "EUSR"
  | "NRSWA"
  | "In-House";

export type NormalizedTrainingOutcome = "Pass" | "Fail" | null;

export interface NormalizedRegisterRecord {
  id: string;
  source: RegisterSource;
  registerKey: AdminRegisterKey;
  candidateName: string;
  companyName: string;
  candidateLookupId: string | null;
  companyLookupId: string | null;
  trainingOutcome: NormalizedTrainingOutcome;
  expiry: string | null;
  /** NPORS category codes e.g. N001 */
  nporsCategories: string[];
  eusrCategory: string | null;
  streetworksCategory: string | null;
  certificateCategory: string | null;
  courseCategory: string | null;
  customerVisible: boolean;
  modifiedAt: string | null;
}

const NPORS_CODES = [
  "N001",
  "N003",
  "N004",
  "N010",
  "N020",
  "N021",
  "N027",
  "N100",
] as const;

export function registerKeyToSource(key: AdminRegisterKey): RegisterSource {
  switch (key) {
    case "nporsRegister":
      return "NPORS";
    case "eusrRegister":
      return "EUSR";
    case "nrswaRegister":
      return "NRSWA";
    case "inHouseCertificates":
      return "In-House";
  }
}

export function sourceToRegisterKey(source: RegisterSource): AdminRegisterKey {
  switch (source) {
    case "NPORS":
      return "nporsRegister";
    case "EUSR":
      return "eusrRegister";
    case "NRSWA":
      return "nrswaRegister";
    case "In-House":
      return "inHouseCertificates";
  }
}

export function normalizeTrainingOutcome(
  value: string | null | undefined,
): NormalizedTrainingOutcome {
  if (!value?.trim()) return null;
  const lower = value.trim().toLowerCase();
  if (lower.startsWith("p")) return "Pass";
  if (lower.startsWith("f")) return "Fail";
  return null;
}

function extractLookupId(
  fields: SharePointFields,
  fieldInternalName: string,
): string | null {
  const direct = fields[fieldInternalName];
  if (direct && typeof direct === "object" && "LookupId" in direct) {
    const id = (direct as { LookupId?: unknown }).LookupId;
    if (typeof id === "number" || typeof id === "string") {
      return String(id);
    }
  }
  const lookupIdField = fields[`${fieldInternalName}LookupId`];
  if (typeof lookupIdField === "number" || typeof lookupIdField === "string") {
    const text = String(lookupIdField).trim();
    return text || null;
  }
  return null;
}

/** Parses MultiChoice / delimited NPORS category values into matrix codes. */
export function parseNporsCategories(value: unknown): string[] {
  const tokens: string[] = [];
  if (Array.isArray(value)) {
    for (const entry of value) {
      if (typeof entry === "string") tokens.push(entry);
    }
  } else if (typeof value === "string" && value.trim()) {
    tokens.push(
      ...value.split(/[;,#|]+/).map((part) => part.trim()).filter(Boolean),
    );
  }

  const found = new Set<string>();
  for (const token of tokens) {
    const upper = token.toUpperCase();
    for (const code of NPORS_CODES) {
      if (upper === code || upper.includes(code)) {
        found.add(code);
      }
    }
  }
  return [...found];
}

/**
 * Graph often returns only CandidateNameLookupId / CompanyNameLookupId.
 * Keep the row when LookupIds exist so matrix sync can still resolve people.
 */
function mapRegisterPeople(
  fields: SharePointFields,
  candidateField: string,
  companyField: string,
): {
  candidateName: string;
  companyName: string;
  candidateLookupId: string | null;
  companyLookupId: string | null;
} | null {
  const candidateLookupId = extractLookupId(fields, candidateField);
  const companyLookupId = extractLookupId(fields, companyField);
  const candidateName =
    asLookupOrString(fields[candidateField]) ??
    asString(fields[candidateField]) ??
    (candidateLookupId ? `Candidate #${candidateLookupId}` : null);
  const companyName =
    asLookupOrString(fields[companyField]) ??
    asString(fields[companyField]) ??
    (companyLookupId ? `Company #${companyLookupId}` : "");

  if (!candidateName) return null;
  if (!companyName && !companyLookupId && !candidateLookupId) return null;

  return {
    candidateName,
    companyName: companyName || "",
    candidateLookupId,
    companyLookupId,
  };
}

function mapRegisterItem(
  key: AdminRegisterKey,
  item: SharePointListItem,
): NormalizedRegisterRecord | null {
  const source = registerKeyToSource(key);

  if (key === "nporsRegister") {
    const f = getSharePointFields("nporsRegister");
    const people = mapRegisterPeople(
      item.fields,
      f.candidateName,
      f.companyName,
    );
    if (!people) return null;
    return {
      id: item.id,
      source,
      registerKey: key,
      candidateName: people.candidateName,
      companyName: people.companyName,
      candidateLookupId: people.candidateLookupId,
      companyLookupId: people.companyLookupId,
      trainingOutcome: normalizeTrainingOutcome(
        asNullableString(item.fields[f.trainingOutcome]),
      ),
      expiry: asNullableString(item.fields[f.expiry]),
      nporsCategories: parseNporsCategories(item.fields[f.nporsCategory]),
      eusrCategory: null,
      streetworksCategory: null,
      certificateCategory: null,
      courseCategory: null,
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      modifiedAt: item.lastModifiedDateTime ?? null,
    };
  }

  if (key === "eusrRegister") {
    const f = getSharePointFields("eusrRegister");
    const people = mapRegisterPeople(
      item.fields,
      f.candidateName,
      f.companyName,
    );
    if (!people) return null;
    return {
      id: item.id,
      source,
      registerKey: key,
      candidateName: people.candidateName,
      companyName: people.companyName,
      candidateLookupId: people.candidateLookupId,
      companyLookupId: people.companyLookupId,
      trainingOutcome: normalizeTrainingOutcome(
        asNullableString(item.fields[f.trainingOutcome]),
      ),
      expiry: asNullableString(item.fields[f.expiry]),
      nporsCategories: [],
      eusrCategory: asNullableString(item.fields[f.eusrCategory]),
      streetworksCategory: null,
      certificateCategory: null,
      courseCategory: null,
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      modifiedAt: item.lastModifiedDateTime ?? null,
    };
  }

  if (key === "nrswaRegister") {
    const f = getSharePointFields("nrswaRegister");
    const people = mapRegisterPeople(
      item.fields,
      f.candidateName,
      f.companyName,
    );
    if (!people) return null;
    return {
      id: item.id,
      source,
      registerKey: key,
      candidateName: people.candidateName,
      companyName: people.companyName,
      candidateLookupId: people.candidateLookupId,
      companyLookupId: people.companyLookupId,
      trainingOutcome: normalizeTrainingOutcome(
        asNullableString(item.fields[f.trainingOutcome]),
      ),
      expiry: asNullableString(item.fields[f.expiryDate]),
      nporsCategories: [],
      eusrCategory: null,
      streetworksCategory: asNullableString(
        item.fields[f.streetworksCategory],
      ),
      certificateCategory: null,
      courseCategory: asNullableString(item.fields[f.course]),
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      modifiedAt: item.lastModifiedDateTime ?? null,
    };
  }

  const f = getSharePointFields("inHouseCertificates");
  const people = mapRegisterPeople(item.fields, f.candidateName, f.companyName);
  if (!people) return null;
  return {
    id: item.id,
    source,
    registerKey: key,
    candidateName: people.candidateName,
    companyName: people.companyName,
    candidateLookupId: people.candidateLookupId,
    companyLookupId: people.companyLookupId,
    trainingOutcome: normalizeTrainingOutcome(
      asNullableString(item.fields[f.trainingOutcome]),
    ),
    expiry: asNullableString(item.fields[f.expiryDate]),
    nporsCategories: [],
    eusrCategory: null,
    streetworksCategory: null,
    certificateCategory: asNullableString(item.fields[f.certificateCategory]),
    courseCategory: asNullableString(item.fields[f.courseCategory]),
    customerVisible: asBoolean(item.fields[f.customerVisible]),
    modifiedAt: item.lastModifiedDateTime ?? null,
  };
}

export async function listNormalizedRegisters(
  key: AdminRegisterKey,
): Promise<NormalizedRegisterRecord[]> {
  const items = await getListItemsByKey(key, { top: 5000 });
  return items
    .map((item) => mapRegisterItem(key, item))
    .filter((row): row is NormalizedRegisterRecord => Boolean(row));
}

export async function listAllNormalizedRegisters(): Promise<
  NormalizedRegisterRecord[]
> {
  // Include In-House so Asbestos Awareness can update N031 on the matrix.
  // Other In-House categories remain ignored in the sync mapper.
  const keys: AdminRegisterKey[] = [
    "nporsRegister",
    "eusrRegister",
    "nrswaRegister",
    "inHouseCertificates",
  ];
  const batches = await Promise.all(keys.map((key) => listNormalizedRegisters(key)));
  return batches.flat();
}

export function normalizeRegisterFromAdminRecord(
  key: AdminRegisterKey,
  record: {
    id: string;
    candidateName: string;
    companyName: string;
    trainingOutcome?: string | null;
    expiry?: string | null;
    nporsCategory?: string | null;
    eusrCategory?: string | null;
    streetworksCategory?: string | null;
    certificateCategory?: string | null;
    courseCategory?: string | null;
    customerVisible?: boolean;
  },
  lookupHints?: {
    candidateLookupId?: string | null;
    companyLookupId?: string | null;
  },
): NormalizedRegisterRecord {
  return {
    id: record.id,
    source: registerKeyToSource(key),
    registerKey: key,
    candidateName: record.candidateName,
    companyName: record.companyName,
    candidateLookupId: lookupHints?.candidateLookupId ?? null,
    companyLookupId: lookupHints?.companyLookupId ?? null,
    trainingOutcome: normalizeTrainingOutcome(record.trainingOutcome),
    expiry: record.expiry ?? null,
    nporsCategories: parseNporsCategories(record.nporsCategory),
    eusrCategory: record.eusrCategory ?? null,
    streetworksCategory: record.streetworksCategory ?? null,
    certificateCategory: record.certificateCategory ?? null,
    courseCategory: record.courseCategory ?? null,
    customerVisible: record.customerVisible ?? true,
    modifiedAt: new Date().toISOString(),
  };
}

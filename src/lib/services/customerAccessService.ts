import "server-only";

import { cache } from "react";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asString,
  getListItemByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import { getWorkforceByCompanyName } from "@/lib/services/workforceService";
import type {
  CustomerContext,
  NormalizedAccessScope,
  WorkforceCandidate,
} from "@/types/models";

const workforceFields = getSharePointFields("workforce");

export function isCompanyWideScope(scope: NormalizedAccessScope): boolean {
  return scope === "Company" || scope === "All";
}

export function accessScopeBadgeLabel(context: CustomerContext): string {
  const scope = context.normalizedAccessScope;
  if (scope === "Company" || scope === "All") return "Company-wide";
  if (scope === "Department") {
    if (context.departmentScopes.length > 0) {
      return `${context.departmentScopes.join(", ")} department`;
    }
    return "Department";
  }
  if (scope === "AssignedCandidates") return "Assigned candidates";
  if (scope === "CandidateOnly") return "Own records only";
  return context.accessScope || "Company";
}

function cell(value: unknown): string {
  return asString(value) ?? "";
}

function lookupDisplay(value: unknown): string | null {
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed || /^\d+$/.test(trimmed)) return null;
    return trimmed;
  }
  if (value && typeof value === "object") {
    return (
      asString((value as { LookupValue?: unknown }).LookupValue) ||
      asString((value as { Title?: unknown }).Title) ||
      asString((value as { CandidateName?: unknown }).CandidateName) ||
      null
    );
  }
  return null;
}

function lookupId(value: unknown): string | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(value);
  }
  if (typeof value === "string" && /^\d+$/.test(value.trim())) {
    return value.trim();
  }
  if (value && typeof value === "object") {
    const id =
      asString((value as { LookupId?: unknown }).LookupId) ||
      asString((value as { Id?: unknown }).Id) ||
      asString((value as { id?: unknown }).id);
    if (id && /^\d+$/.test(id)) return id;
  }
  return null;
}

function splitDepartments(value: string | null | undefined): string[] {
  if (!value?.trim()) return [];
  return value
    .split(/;|,/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function namesMatch(
  left: string | null | undefined,
  right: string | null | undefined,
): boolean {
  const a = (left ?? "").trim().toLowerCase();
  const b = (right ?? "").trim().toLowerCase();
  return Boolean(a && b && a === b);
}

function personMatchesContext(
  person: string | null | undefined,
  context: CustomerContext,
): boolean {
  if (namesMatch(person, context.candidateScopeName)) return true;
  if (namesMatch(person, context.loggedInEmail)) return true;
  if (namesMatch(person, context.roleLabel)) return true;
  // Permissions "Name" often matches Workforce Training Manager / Supervisor text.
  const emailLocal = context.loggedInEmail.split("@")[0]?.trim();
  if (emailLocal && namesMatch(person, emailLocal)) return true;
  return false;
}

function supervisorMatchesCandidate(
  candidate: WorkforceCandidate,
  context: CustomerContext,
): boolean {
  return personMatchesContext(candidate.supervisor, context);
}

function trainingManagerMatchesCandidate(
  candidate: Pick<WorkforceCandidate, "trainingManager">,
  context: CustomerContext,
): boolean {
  return personMatchesContext(candidate.trainingManager, context);
}

/**
 * Whether a workforce / matrix-style row is visible under the user's access scope.
 * Company-wide roles see everything in the already company-filtered set.
 */
export function candidateRecordAllowed(
  candidate: Pick<
    WorkforceCandidate,
    "candidateName" | "department" | "supervisor" | "trainingManager"
  > & { email?: string | null },
  context: CustomerContext,
): boolean {
  if (isCompanyWideScope(context.normalizedAccessScope)) {
    return true;
  }

  if (context.normalizedAccessScope === "CandidateOnly") {
    const email = candidate.email?.trim().toLowerCase();
    if (email && email === context.loggedInEmail) return true;
    if (
      context.candidateScopeName &&
      namesMatch(candidate.candidateName, context.candidateScopeName)
    ) {
      return true;
    }
    return false;
  }

  // Department / AssignedCandidates — match department coverage first.
  if (context.departmentScopes.length > 0) {
    const depts = splitDepartments(candidate.department);
    if (
      depts.some((dept) =>
        context.departmentScopes.some((scope) => namesMatch(dept, scope)),
      )
    ) {
      return true;
    }
  }

  // Training Managers see candidates assigned to them on Workforce.
  if (trainingManagerMatchesCandidate(candidate, context)) {
    return true;
  }

  // Supervisors can also see candidates assigned to them by name.
  if (supervisorMatchesCandidate(candidate as WorkforceCandidate, context)) {
    return true;
  }

  // Department-only with no scopes and no assignment match → nothing visible.
  return false;
}

export function filterCandidatesByAccess(
  candidates: WorkforceCandidate[],
  context: CustomerContext,
): WorkforceCandidate[] {
  if (isCompanyWideScope(context.normalizedAccessScope)) {
    return candidates;
  }
  return candidates.filter((row) => candidateRecordAllowed(row, context));
}

/** Filter rows that expose a candidate name (+ optional department). */
export function filterRowsByCandidateAccess<
  T extends { candidateName: string; department?: string | null },
>(rows: T[], allowedNames: Set<string>, context: CustomerContext): T[] {
  if (isCompanyWideScope(context.normalizedAccessScope)) {
    return rows;
  }

  return rows.filter((row) => {
    const name = row.candidateName.trim().toLowerCase();
    if (allowedNames.has(name)) return true;

    if (
      context.normalizedAccessScope === "Department" &&
      context.departmentScopes.length > 0 &&
      row.department
    ) {
      return splitDepartments(row.department).some((dept) =>
        context.departmentScopes.some((scope) => namesMatch(dept, scope)),
      );
    }

    if (
      context.normalizedAccessScope === "CandidateOnly" &&
      context.candidateScopeName
    ) {
      return namesMatch(row.candidateName, context.candidateScopeName);
    }

    return false;
  });
}

/**
 * Deduped per request — layout/pages/APIs often need workforce more than once.
 */
export const getAllowedWorkforceForCustomer = cache(
  async (context: CustomerContext): Promise<WorkforceCandidate[]> => {
    const all = await getWorkforceByCompanyName(context.companyName);
    return filterCandidatesByAccess(all, context);
  },
);

export async function getAllowedCandidateNames(
  context: CustomerContext,
): Promise<Set<string>> {
  const allowed = await getAllowedWorkforceForCustomer(context);
  return new Set(
    allowed.map((row) => row.candidateName.trim().toLowerCase()).filter(Boolean),
  );
}

export async function getAllowedCandidateIds(
  context: CustomerContext,
): Promise<Set<string>> {
  const allowed = await getAllowedWorkforceForCustomer(context);
  return new Set(allowed.map((row) => row.id));
}

/** Full company workforce id → name map (one Graph list read). */
export async function getCompanyCandidateNameMap(
  companyName: string,
): Promise<Map<string, string>> {
  const workforce = await getWorkforceByCompanyName(companyName);
  const map = new Map<string, string>();
  for (const row of workforce) {
    map.set(row.id, row.candidateName);
  }
  return map;
}

export function assertCandidateAccess(
  candidate: WorkforceCandidate,
  context: CustomerContext,
): boolean {
  return (
    namesMatch(candidate.companyName, context.companyName) &&
    candidateRecordAllowed(candidate, context)
  );
}

/** Sync resolve using an already-loaded name map — avoids per-document Graph calls. */
export function resolveCandidateDisplayNameSync(
  candidateValue: unknown,
  candidateIdHint: string | null,
  nameCache: Map<string, string>,
): string | null {
  const display = lookupDisplay(candidateValue);
  if (display) return display;

  const id = candidateIdHint || lookupId(candidateValue);
  if (!id) {
    const raw = cell(candidateValue);
    if (raw && !/^\d+$/.test(raw)) return raw;
    return null;
  }

  const cached = nameCache.get(id);
  if (cached) return cached;
  return null;
}

/** Resolve Candidate lookup / id to a display name (never return raw numeric IDs). */
export async function resolveCandidateDisplayName(
  candidateValue: unknown,
  candidateIdHint: string | null,
  nameCache: Map<string, string>,
): Promise<string | null> {
  const sync = resolveCandidateDisplayNameSync(
    candidateValue,
    candidateIdHint,
    nameCache,
  );
  if (sync) return sync;

  const id = candidateIdHint || lookupId(candidateValue);
  if (!id) return null;
  if (nameCache.has(id)) {
    const cached = nameCache.get(id);
    return cached || null;
  }

  try {
    const item = await getListItemByKey("workforce", id);
    const name =
      (item && cell(item.fields[workforceFields.candidateName])) ||
      (item && cell(item.fields.Title)) ||
      "";
    const resolved = name || null;
    nameCache.set(id, resolved ?? "");
    return resolved;
  } catch {
    nameCache.set(id, "");
    return null;
  }
}

export function extractCandidateLookupId(
  fields: SharePointFields,
  candidateField: string,
): string | null {
  return (
    lookupId(fields[candidateField]) ||
    cell(fields.CandidateId) ||
    cell(fields[`${candidateField}LookupId`]) ||
    null
  );
}

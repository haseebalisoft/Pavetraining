import { getSharePointFields } from "../schema/sharepointSchema";
import type { PermissionProfile } from "../types/models";
import { accessScopeBadgeLabel } from "./permissionService";
import {
  asBoolean,
  asString,
  getListItem,
  getListItems,
  normalizeSharePointUserEmail,
  type SpListClient,
} from "./sharePointListService";
import type { PortalTableRow } from "./portalDataService";

const docs = getSharePointFields("customerDocuments");
const workforce = getSharePointFields("workforce");

export interface CustomerDocumentRow {
  id: string;
  name: string;
  documentType: string;
  candidateName: string;
  companyName: string;
  modifiedDate: string;
  viewUrl: string | null;
  downloadUrl: string | null;
  canDownload: boolean;
}

function cell(value: unknown): string {
  return asString(value) ?? "";
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

function isSharePointFile(fields: Record<string, unknown>): boolean {
  const fs = fields[docs.fsObjType];
  if (fs === 1 || fs === "1") return false;
  if (fs === 0 || fs === "0") return true;
  const leaf = cell(fields[docs.fileLeafRef]);
  if (!leaf) return false;
  return leaf.indexOf(".") >= 0;
}

function formatDate(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

function departmentsFromWorkforce(fields: Record<string, unknown>): string[] {
  const raw = fields[workforce.department];
  if (Array.isArray(raw)) {
    return raw.map((x) => cell(x)).filter(Boolean);
  }
  const single = cell(raw);
  if (!single) return [];
  return single.split(/;|,/).map((p) => p.trim()).filter(Boolean);
}

function supervisorMatches(
  fields: Record<string, unknown>,
  permission: PermissionProfile
): boolean {
  const supervisor = fields[workforce.supervisor];
  const display = lookupDisplay(supervisor);
  const id = lookupId(supervisor) || cell(fields.SupervisorId);

  if (id && id === permission.permissionItemId) return true;
  if (
    display &&
    permission.candidateScopeName &&
    display.toLowerCase() === permission.candidateScopeName.toLowerCase()
  ) {
    return true;
  }
  if (display && display.toLowerCase() === permission.userEmail.toLowerCase()) {
    return true;
  }
  if (
    display &&
    permission.roleLabel &&
    display.toLowerCase() === permission.roleLabel.toLowerCase()
  ) {
    return true;
  }
  return false;
}

/**
 * Whether a workforce / matrix-style row is visible under the user's access scope.
 */
export function candidateRecordAllowed(
  fields: Record<string, unknown>,
  permission: PermissionProfile,
  options?: { candidateNameField?: string; departmentField?: string }
): boolean {
  const scope = permission.normalizedAccessScope;
  if (scope === "Company" || scope === "All") return true;

  const nameField = options?.candidateNameField || workforce.candidateName;
  const deptField = options?.departmentField || workforce.department;
  const candidateName = cell(fields[nameField]);

  if (scope === "CandidateOnly") {
    const email = normalizeSharePointUserEmail(cell(fields[workforce.email]));
    if (email && email === permission.userEmail) return true;
    if (
      permission.candidateScopeName &&
      candidateName &&
      candidateName.toLowerCase() ===
        permission.candidateScopeName.toLowerCase()
    ) {
      return true;
    }
    return false;
  }

  // Department / AssignedCandidates
  if (supervisorMatches(fields, permission)) return true;

  if (permission.departmentScopes.length > 0) {
    const depts = departmentsFromWorkforce({
      ...fields,
      [workforce.department]: fields[deptField],
    });
    for (let i = 0; i < depts.length; i++) {
      for (let j = 0; j < permission.departmentScopes.length; j++) {
        if (
          depts[i].toLowerCase() ===
          permission.departmentScopes[j].toLowerCase()
        ) {
          return true;
        }
      }
    }
    return false;
  }

  // No department scopes configured — fall back to supervisor assignment only.
  return supervisorMatches(fields, permission);
}

export function filterPortalRowsByAccess(
  rows: PortalTableRow[],
  permission: PermissionProfile,
  options?: { candidateNameField?: string; departmentField?: string }
): PortalTableRow[] {
  if (
    permission.normalizedAccessScope === "Company" ||
    permission.normalizedAccessScope === "All"
  ) {
    return rows;
  }
  return rows.filter((row) =>
    candidateRecordAllowed(row.fields || {}, permission, options)
  );
}

async function resolveCandidateName(
  client: SpListClient,
  candidateValue: unknown,
  candidateIdHint: string | null,
  cache: Record<string, string>
): Promise<string> {
  const display = lookupDisplay(candidateValue);
  if (display) return display;

  const id = candidateIdHint || lookupId(candidateValue);
  if (!id) return "—";
  if (cache[id]) return cache[id];

  try {
    const item = await getListItem(client, "workforce", id);
    const name =
      (item && cell(item.fields[workforce.candidateName])) ||
      (item && cell(item.fields.Title)) ||
      "";
    cache[id] = name || "—";
    return cache[id];
  } catch {
    cache[id] = "—";
    return "—";
  }
}

function buildFileUrl(webUrl: string, fields: Record<string, unknown>): string | null {
  const abs = cell(fields.EncodedAbsUrl);
  if (abs) return abs;
  const fileRef = cell(fields[docs.fileRef]);
  if (!fileRef) return null;
  if (/^https?:\/\//i.test(fileRef)) return fileRef;
  const root = webUrl.replace(/\/$/, "");
  if (fileRef.indexOf("/") === 0) return root.replace(/^(https?:\/\/[^/]+).*$/i, "$1") + fileRef;
  return root + "/" + fileRef;
}

/**
 * Customer Documents — company + CustomerVisible + files only,
 * with candidate names resolved and role scope applied.
 */
export async function loadCustomerDocuments(
  client: SpListClient,
  permission: PermissionProfile
): Promise<CustomerDocumentRow[]> {
  const companyId = permission.companyId;
  const companyName = permission.companyDisplayName || "";
  if (!companyId || companyId === "0") return [];

  const items = await getListItems(client, "customerDocuments", {
    filter: ["CompanyId eq " + Number(companyId), docs.customerVisible + " eq 1"].join(
      " and "
    ),
    top: 5000,
    maxItems: 5000,
  });

  const nameCache: Record<string, string> = {};
  const allowedCandidateIds = new Set<string>();
  let workforceIndex: Array<{ id: string; fields: Record<string, unknown> }> | null =
    null;

  const needsScopeFilter =
    permission.normalizedAccessScope !== "Company" &&
    permission.normalizedAccessScope !== "All";

  if (needsScopeFilter) {
    workforceIndex = await getListItems(client, "workforce", {
      filter: "CompanyNameId eq " + Number(companyId),
      top: 5000,
      maxItems: 5000,
    }).catch(async () =>
      getListItems(client, "workforce", {
        top: 5000,
        maxItems: 5000,
      })
    );

    for (let i = 0; i < workforceIndex.length; i++) {
      const wf = workforceIndex[i];
      if (candidateRecordAllowed(wf.fields, permission)) {
        allowedCandidateIds.add(wf.id);
        const n = cell(wf.fields[workforce.candidateName]);
        if (n) nameCache[wf.id] = n;
      }
    }
  }

  const rows: CustomerDocumentRow[] = [];

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const fields = item.fields;
    if (!isSharePointFile(fields)) continue;
    if (!asBoolean(fields[docs.customerVisible])) continue;

    const companyDisplay =
      lookupDisplay(fields[docs.company]) ||
      cell(fields.Company) ||
      companyName;
    if (!companyDisplay && !companyId) continue;

    const candidateRaw = fields[docs.candidate];
    const candidateId =
      lookupId(candidateRaw) || cell(fields.CandidateId) || null;

    if (needsScopeFilter) {
      // Company-level docs (no candidate) — Training Manager only; supervisors skip.
      if (!candidateId) {
        if (permission.customerRole !== "TrainingManager") continue;
      } else if (!allowedCandidateIds.has(candidateId)) {
        continue;
      }
    }

    const candidateName = await resolveCandidateName(
      client,
      candidateRaw,
      candidateId,
      nameCache
    );

    const name =
      cell(fields[docs.fileLeafRef]) ||
      cell(fields[docs.title]) ||
      "Document";
    const modified =
      cell(fields[docs.modified]) ||
      cell(fields.Modified) ||
      "";
    const viewUrl = buildFileUrl(client.webUrl, fields);
    const canDownload = permission.canDownload === true;

    rows.push({
      id: item.id,
      name,
      documentType: cell(fields[docs.documentType]) || "Document",
      candidateName,
      companyName: companyDisplay || companyName,
      modifiedDate: formatDate(modified),
      viewUrl,
      downloadUrl: canDownload && viewUrl ? viewUrl : null,
      canDownload,
    });
  }

  return rows;
}

export function documentRowsToPortalTable(
  docsRows: CustomerDocumentRow[]
): PortalTableRow[] {
  return docsRows.map((row) => ({
    id: row.id,
    cells: [
      row.name,
      row.documentType,
      row.candidateName,
      row.modifiedDate,
      row.viewUrl || "",
      row.canDownload && row.downloadUrl ? row.downloadUrl : "",
    ],
    fields: {
      __docViewUrl: row.viewUrl,
      __docDownloadUrl: row.downloadUrl,
      __docCanDownload: row.canDownload,
    },
  }));
}

export { accessScopeBadgeLabel };

import "server-only";

import {
  getSharePointListId,
  getSharePointSiteApiRoot,
} from "@/lib/config/sharepoint";
import { getGraphClient } from "@/lib/graph/graphClient";
import {
  CANDIDATE_SUBFOLDERS,
  COMPANY_LEVEL_FOLDERS,
  candidateDocumentsFolderName,
  companyDocumentsFolderName,
  resolveDocumentFolderSegments,
  resolveDocumentTypeFolder,
  sanitizeFolderSegment,
  type CandidateDocumentSubfolder,
  type DocumentDestinationFolder,
  type DocumentFolderResolveInput,
  type DocumentFolderSegments,
} from "@/lib/services/documentFolderPaths";

// Folder naming + path rules live in the dependency-free documentFolderPaths
// module (single source of truth, unit-testable). Re-export so existing
// importers of this service keep working unchanged.
export {
  CANDIDATE_SUBFOLDERS,
  COMPANY_LEVEL_FOLDERS,
  candidateDocumentsFolderName,
  companyDocumentsFolderName,
  resolveDocumentFolderSegments,
  resolveDocumentTypeFolder,
  sanitizeFolderSegment,
};
export type {
  CandidateDocumentSubfolder,
  DocumentDestinationFolder,
  DocumentFolderResolveInput,
  DocumentFolderSegments,
};

export interface ResolvedDocumentUploadFolder {
  driveId: string;
  folderId: string;
  segments: DocumentFolderSegments;
  destinationFolder: DocumentDestinationFolder;
  /** Display path joined with `/`. */
  path: string;
}

async function getCustomerDocumentsDriveId(): Promise<string> {
  const siteRoot = getSharePointSiteApiRoot();
  const listId = getSharePointListId("customerDocuments");
  const client = getGraphClient();
  const drive = (await client
    .api(`${siteRoot}/lists/${listId}/drive`)
    .get()) as { id?: string };
  if (!drive.id) {
    throw new Error("Customer Documents drive could not be resolved.");
  }
  return drive.id;
}

export type CustomerDocumentsDriveChild = {
  driveItemId: string;
  name: string;
  isFolder: boolean;
  lastModifiedDateTime: string | null;
  listItemId: string | null;
  fields: Record<string, unknown>;
  webUrl: string | null;
};

function encodeDrivePath(segments: string[]): string {
  return segments
    .map((segment) => encodeURIComponent(sanitizeFolderSegment(segment)))
    .filter(Boolean)
    .join("/");
}

/**
 * Lists immediate children of a Customer Documents folder path
 * (SharePoint drive browse — same hierarchy as the library UI).
 *
 * pathSegments example:
 * [] → root company folders
 * ["C00024 - Murphy Plant Ltd"] → Company Documents + Candidates
 * ["C00024 - Murphy Plant Ltd", "Candidates"] → candidate folders
 * ["C00024 - Murphy Plant Ltd", "Candidates", "W00195 - John Murphy Test"]
 *   → Certificates | Card Scans | NVQ Documents | Other Documents
 */
export async function browseCustomerDocumentsFolder(
  pathSegments: string[],
): Promise<CustomerDocumentsDriveChild[]> {
  const driveId = await getCustomerDocumentsDriveId();
  const client = getGraphClient();
  const safeSegments = pathSegments
    .map((segment) => sanitizeFolderSegment(segment))
    .filter(Boolean);

  const apiPath =
    safeSegments.length === 0
      ? `/drives/${driveId}/root/children`
      : `/drives/${driveId}/root:/${encodeDrivePath(safeSegments)}:/children`;

  const response = (await client
    .api(apiPath)
    .expand("listItem($expand=fields)")
    .top(200)
    .get()) as {
    value?: Array<{
      id?: string;
      name?: string;
      folder?: unknown;
      file?: unknown;
      lastModifiedDateTime?: string;
      webUrl?: string;
      listItem?: { id?: string; fields?: Record<string, unknown> };
    }>;
  };

  return (response.value ?? [])
    .filter((item) => Boolean(item.id && item.name))
    .map((item) => ({
      driveItemId: String(item.id),
      name: String(item.name),
      isFolder: Boolean(item.folder),
      lastModifiedDateTime: item.lastModifiedDateTime ?? null,
      listItemId: item.listItem?.id ? String(item.listItem.id) : null,
      fields: item.listItem?.fields ?? {},
      webUrl: item.webUrl ?? null,
    }))
    .sort((a, b) => {
      if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
      return a.name.localeCompare(b.name);
    });
}

async function listChildFolders(
  driveId: string,
  parentId: string,
): Promise<Array<{ id: string; name: string }>> {
  const client = getGraphClient();
  const response = (await client
    .api(`/drives/${driveId}/items/${parentId}/children`)
    .filter("folder ne null")
    .top(999)
    .get()) as { value?: Array<{ id?: string; name?: string }> };

  return (response.value ?? [])
    .filter((item): item is { id: string; name: string } =>
      Boolean(item.id && item.name),
    )
    .map((item) => ({ id: String(item.id), name: String(item.name) }));
}

/**
 * Prefer exact folder name; if a stable number is provided, reuse any folder
 * whose name starts with `{number} -` so renames do not create duplicates.
 */
async function resolveChildFolderId(
  driveId: string,
  parentId: string,
  preferredName: string,
  stableNumber?: string | null,
): Promise<string | null> {
  const children = await listChildFolders(driveId, parentId);
  const preferred = preferredName.toLowerCase();
  const exact = children.find((item) => item.name.toLowerCase() === preferred);
  if (exact) {
    return exact.id;
  }

  const number = sanitizeFolderSegment(stableNumber ?? "");
  if (!number) {
    return null;
  }

  const prefix = `${number.toLowerCase()} -`;
  const byNumber = children.find((item) => {
    const lower = item.name.toLowerCase();
    return lower === number.toLowerCase() || lower.startsWith(prefix);
  });
  return byNumber?.id ?? null;
}

async function ensureChildFolder(
  driveId: string,
  parentId: string,
  name: string,
  stableNumber?: string | null,
): Promise<string> {
  const existing = await resolveChildFolderId(
    driveId,
    parentId,
    name,
    stableNumber,
  );
  if (existing) {
    return existing;
  }

  const client = getGraphClient();
  try {
    const created = (await client
      .api(`/drives/${driveId}/items/${parentId}/children`)
      .post({
        name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      })) as { id?: string };
    if (!created.id) {
      throw new Error(`Folder "${name}" was created without an id.`);
    }
    return String(created.id);
  } catch (error) {
    const again = await resolveChildFolderId(
      driveId,
      parentId,
      name,
      stableNumber,
    );
    if (again) return again;
    throw error;
  }
}

type PathStep = {
  name: string;
  /** When set, match existing folders by this number prefix. */
  stableNumber?: string | null;
};

/**
 * Memoizes the drive id + resolved folder ids for the duration of ONE bulk
 * import. Without it, every candidate re-fetched the drive id (~6×/candidate)
 * and re-resolved the shared company / "Candidates" folders from root for every
 * subfolder — the dominant cost of a 50-row import. Values are stored as
 * in-flight PROMISES so parallel candidates of the same company share a single
 * resolve instead of racing to create the same folder.
 */
export interface FolderEnsureCache {
  driveId?: Promise<string>;
  folders: Map<string, Promise<string>>;
}

export function createFolderEnsureCache(): FolderEnsureCache {
  return { folders: new Map() };
}

function getDriveIdCached(cache: FolderEnsureCache): Promise<string> {
  if (!cache.driveId) {
    const pending = getCustomerDocumentsDriveId();
    cache.driveId = pending;
    // On failure, clear so a later attempt can retry.
    pending.catch(() => {
      if (cache.driveId === pending) cache.driveId = undefined;
    });
  }
  return cache.driveId;
}

function ensureChildFolderCached(
  cache: FolderEnsureCache,
  driveId: string,
  parentId: string,
  name: string,
  stableNumber?: string | null,
): Promise<string> {
  const key = `${driveId}::${parentId}::${name.toLowerCase()}`;
  const hit = cache.folders.get(key);
  if (hit) return hit;
  const pending = ensureChildFolder(driveId, parentId, name, stableNumber);
  cache.folders.set(key, pending);
  pending.catch(() => {
    if (cache.folders.get(key) === pending) cache.folders.delete(key);
  });
  return pending;
}

async function ensureFolderPathStepsCached(
  steps: PathStep[],
  cache: FolderEnsureCache,
): Promise<{ driveId: string; folderId: string }> {
  const driveId = await getDriveIdCached(cache);
  let parentId = "root";
  for (const step of steps) {
    const safe = sanitizeFolderSegment(step.name);
    if (!safe) continue;
    parentId = await ensureChildFolderCached(
      cache,
      driveId,
      parentId,
      safe,
      step.stableNumber,
    );
  }
  return { driveId, folderId: parentId };
}

async function ensureFolderPathSteps(steps: PathStep[]): Promise<{
  driveId: string;
  folderId: string;
}> {
  // Single-call callers get a throwaway cache: still fetches the drive id once
  // and dedupes repeated segments within the one call.
  return ensureFolderPathStepsCached(steps, createFolderEnsureCache());
}

function companyFolderStep(
  companyNumber: string | null | undefined,
  companyName: string,
): PathStep {
  return {
    name: companyDocumentsFolderName(companyNumber, companyName),
    stableNumber: companyNumber,
  };
}

function candidateFolderStep(
  workforceNumber: string | null | undefined,
  candidateName: string,
): PathStep {
  return {
    name: candidateDocumentsFolderName(workforceNumber, candidateName),
    stableNumber: workforceNumber,
  };
}

/**
 * Ensures the destination folder for an upload exists (never recreates the
 * library) and returns the drive folder id to upload into.
 */
export async function resolveDocumentUploadFolder(
  input: DocumentFolderResolveInput,
): Promise<ResolvedDocumentUploadFolder> {
  const { segments, destinationFolder } = resolveDocumentFolderSegments(input);
  const hasCandidate = destinationFolder !== "Company Documents";

  const steps: PathStep[] = [
    companyFolderStep(input.companyNumber, input.companyName),
  ];

  if (!hasCandidate) {
    steps.push({ name: "Company Documents" });
  } else {
    steps.push({ name: "Candidates" });
    steps.push(
      candidateFolderStep(
        input.workforceNumber,
        input.candidateName ?? "Candidate",
      ),
    );
    steps.push({ name: destinationFolder });
  }

  const { driveId, folderId } = await ensureFolderPathSteps(steps);

  return {
    driveId,
    folderId,
    segments,
    destinationFolder,
    path: segments.join("/"),
  };
}

/**
 * Ensures:
 * {CompanyNumber - Company Name}/Company Documents
 * {CompanyNumber - Company Name}/Candidates
 *
 * Never throws to callers that should keep list-item create successful —
 * returns a warning string on failure.
 */
export async function ensureCompanyDocumentFolders(
  input: {
    companyName: string;
    companyNumber?: string | null;
  },
  cache: FolderEnsureCache = createFolderEnsureCache(),
): Promise<{ ok: true } | { ok: false; warning: string }> {
  try {
    const company = companyFolderStep(input.companyNumber, input.companyName);
    await ensureFolderPathStepsCached([company, { name: "Company Documents" }], cache);
    await ensureFolderPathStepsCached([company, { name: "Candidates" }], cache);
    return { ok: true };
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Folder create failed.";
    console.warn("[customerDocumentsFolderService] company folders:", message);
    return {
      ok: false,
      warning: `Company saved, but document folders were not created: ${message}`,
    };
  }
}

/**
 * Ensures company folders, then:
 * .../Candidates/{WorkforceNumber - Candidate}/
 *   Certificates | Card Scans | NVQ Documents | Other Documents
 */
export async function ensureCandidateDocumentFolders(
  input: {
    companyName: string;
    companyNumber?: string | null;
    candidateName: string;
    workforceNumber?: string | null;
  },
  cache: FolderEnsureCache = createFolderEnsureCache(),
): Promise<{ ok: true } | { ok: false; warning: string }> {
  try {
    const company = companyFolderStep(input.companyNumber, input.companyName);
    const candidate = candidateFolderStep(
      input.workforceNumber,
      input.candidateName,
    );

    // Company-level folders — memoized, so a batch resolves them once per
    // company instead of once per candidate per subfolder.
    await ensureFolderPathStepsCached([company, { name: "Company Documents" }], cache);
    const { driveId, folderId: candidatesId } = await ensureFolderPathStepsCached(
      [company, { name: "Candidates" }],
      cache,
    );

    // Candidate folder under Candidates, then the four required subfolders in
    // parallel (distinct names under a known parent — safe, and removes the
    // repeated root-to-Candidates walk the loop used to do per subfolder).
    const candidateId = await ensureChildFolderCached(
      cache,
      driveId,
      candidatesId,
      sanitizeFolderSegment(candidate.name),
      candidate.stableNumber,
    );
    await Promise.all(
      CANDIDATE_SUBFOLDERS.map((sub) =>
        ensureChildFolderCached(cache, driveId, candidateId, sub),
      ),
    );
    return { ok: true };
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Folder create failed.";
    console.warn(
      "[customerDocumentsFolderService] candidate folders:",
      message,
    );
    return {
      ok: false,
      warning: `Candidate saved, but document folders were not created: ${message}`,
    };
  }
}

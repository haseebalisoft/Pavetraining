import "server-only";

import {
  getSharePointListId,
  getSharePointSiteApiRoot,
} from "@/lib/config/sharepoint";
import { getGraphClient } from "@/lib/graph/graphClient";

const CANDIDATE_SUBFOLDERS = [
  "Certificates",
  "Card Scans",
  "NVQ Documents",
  "Other Documents",
] as const;

export type CandidateDocumentSubfolder = (typeof CANDIDATE_SUBFOLDERS)[number];

export type DocumentDestinationFolder =
  | "Company Documents"
  | CandidateDocumentSubfolder;

export type DocumentFolderSegments = string[];

export interface DocumentFolderResolveInput {
  companyNumber?: string | null;
  companyName: string;
  workforceNumber?: string | null;
  candidateName?: string | null;
  documentType?: string | null;
  /** When false / omitted with no candidate name, routes to Company Documents. */
  hasCandidate?: boolean;
}

export interface ResolvedDocumentUploadFolder {
  driveId: string;
  folderId: string;
  segments: DocumentFolderSegments;
  destinationFolder: DocumentDestinationFolder;
  /** Display path joined with `/`. */
  path: string;
}

function sanitizeFolderSegment(value: string): string {
  return value
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

export function companyDocumentsFolderName(
  companyNumber: string | null | undefined,
  companyName: string,
): string {
  const name = sanitizeFolderSegment(companyName) || "Company";
  const number = sanitizeFolderSegment(companyNumber ?? "");
  return number ? `${number} - ${name}` : name;
}

export function candidateDocumentsFolderName(
  workforceNumber: string | null | undefined,
  candidateName: string,
): string {
  const name = sanitizeFolderSegment(candidateName) || "Candidate";
  const number = sanitizeFolderSegment(workforceNumber ?? "");
  return number ? `${number} - ${name}` : name;
}

/**
 * Maps portal Document Type (+ whether a candidate is attached) to the
 * physical library folder under Customer Documents.
 */
export function resolveDocumentTypeFolder(options: {
  documentType?: string | null;
  hasCandidate: boolean;
}): DocumentDestinationFolder {
  if (!options.hasCandidate) {
    return "Company Documents";
  }

  const normalized = (options.documentType ?? "").trim().toLowerCase();

  if (
    normalized === "certificate" ||
    normalized === "certificates" ||
    normalized.includes("certificate")
  ) {
    return "Certificates";
  }

  if (
    normalized === "card scan" ||
    normalized === "card scans" ||
    normalized.includes("card scan") ||
    (normalized.includes("card") && normalized.includes("scan"))
  ) {
    return "Card Scans";
  }

  if (
    normalized === "nvq document" ||
    normalized === "nvq documents" ||
    normalized.includes("nvq")
  ) {
    return "NVQ Documents";
  }

  return "Other Documents";
}

/**
 * Builds the expected folder path segments using stable company / workforce
 * numbers (names are display suffixes only).
 */
export function resolveDocumentFolderSegments(
  input: DocumentFolderResolveInput,
): {
  segments: DocumentFolderSegments;
  destinationFolder: DocumentDestinationFolder;
} {
  const hasCandidate = Boolean(
    input.hasCandidate ??
      (input.candidateName?.trim() || input.workforceNumber?.trim()),
  );

  const destinationFolder = resolveDocumentTypeFolder({
    documentType: input.documentType,
    hasCandidate,
  });

  const companyFolder = companyDocumentsFolderName(
    input.companyNumber,
    input.companyName,
  );

  if (!hasCandidate || destinationFolder === "Company Documents") {
    return {
      segments: [companyFolder, "Company Documents"],
      destinationFolder: "Company Documents",
    };
  }

  const candidateFolder = candidateDocumentsFolderName(
    input.workforceNumber,
    input.candidateName ?? "Candidate",
  );

  return {
    segments: [
      companyFolder,
      "Candidates",
      candidateFolder,
      destinationFolder,
    ],
    destinationFolder,
  };
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

/** Expected top-level children under a company folder. */
export const COMPANY_LEVEL_FOLDERS = [
  "Company Documents",
  "Candidates",
] as const;

export { CANDIDATE_SUBFOLDERS };

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

async function ensureFolderPathSteps(steps: PathStep[]): Promise<{
  driveId: string;
  folderId: string;
}> {
  const driveId = await getCustomerDocumentsDriveId();
  let parentId = "root";
  for (const step of steps) {
    const safe = sanitizeFolderSegment(step.name);
    if (!safe) continue;
    parentId = await ensureChildFolder(
      driveId,
      parentId,
      safe,
      step.stableNumber,
    );
  }
  return { driveId, folderId: parentId };
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
export async function ensureCompanyDocumentFolders(input: {
  companyName: string;
  companyNumber?: string | null;
}): Promise<{ ok: true } | { ok: false; warning: string }> {
  try {
    const company = companyFolderStep(input.companyNumber, input.companyName);
    await ensureFolderPathSteps([company, { name: "Company Documents" }]);
    await ensureFolderPathSteps([company, { name: "Candidates" }]);
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
export async function ensureCandidateDocumentFolders(input: {
  companyName: string;
  companyNumber?: string | null;
  candidateName: string;
  workforceNumber?: string | null;
}): Promise<{ ok: true } | { ok: false; warning: string }> {
  try {
    const company = companyFolderStep(input.companyNumber, input.companyName);
    const candidate = candidateFolderStep(
      input.workforceNumber,
      input.candidateName,
    );

    await ensureFolderPathSteps([company, { name: "Company Documents" }]);
    await ensureFolderPathSteps([company, { name: "Candidates" }]);

    for (const sub of CANDIDATE_SUBFOLDERS) {
      await ensureFolderPathSteps([
        company,
        { name: "Candidates" },
        candidate,
        { name: sub },
      ]);
    }
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

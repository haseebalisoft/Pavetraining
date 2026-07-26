import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { getCompanyById } from "@/lib/services/companyService";
import {
  resolveDocumentFolderSegments,
  type DocumentDestinationFolder,
} from "@/lib/services/customerDocumentsFolderService";
import {
  asLookupOrString,
  asNullableString,
  asString,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import { getWorkforceById } from "@/lib/services/workforceService";

const documentFields = getSharePointFields("customerDocuments");

export type DocumentMigrationMismatch = {
  id: string;
  fileName: string;
  company: string | null;
  companyId: string | null;
  candidate: string | null;
  candidateId: string | null;
  documentType: string | null;
  currentFolderPath: string;
  expectedFolderPath: string;
  destinationFolder: DocumentDestinationFolder;
  fileRef: string | null;
  reason: string;
};

export type DocumentMigrationReport = {
  generatedAt: string;
  scannedFiles: number;
  matched: number;
  mismatched: number;
  skipped: number;
  /** Files whose folder path does not match metadata-driven routing. */
  mismatches: DocumentMigrationMismatch[];
  notes: string[];
};

function isSharePointFolder(fields: SharePointFields): boolean {
  const fs = fields[documentFields.fsObjType];
  return fs === 1 || fs === "1";
}

function decodePathSegment(segment: string): string {
  try {
    return decodeURIComponent(segment.replace(/\+/g, " "));
  } catch {
    return segment;
  }
}

/**
 * Extracts folder segments under the Customer Documents library from FileRef.
 */
export function extractCustomerDocumentsFolderPath(
  fileRef: string | null | undefined,
): string[] {
  if (!fileRef?.trim()) return [];

  const parts = fileRef
    .split("/")
    .map((part) => decodePathSegment(part))
    .filter(Boolean);

  const libraryIdx = parts.findIndex(
    (part) => part.toLowerCase() === "customer documents",
  );
  const afterLibrary = libraryIdx >= 0 ? parts.slice(libraryIdx + 1) : parts;
  if (afterLibrary.length <= 1) {
    return [];
  }
  // Drop the file leaf.
  return afterLibrary.slice(0, -1);
}

function pathsEqual(actual: string[], expected: string[]): boolean {
  if (actual.length !== expected.length) return false;
  return actual.every(
    (segment, index) =>
      segment.trim().toLowerCase() === expected[index].trim().toLowerCase(),
  );
}

/**
 * Number-stable path match: company / candidate folders may keep an older
 * display name after a rename, as long as the number prefix matches.
 */
function pathsMatchAllowingNumberStableNames(
  actual: string[],
  expected: string[],
  companyNumber: string | null,
  workforceNumber: string | null,
): boolean {
  if (pathsEqual(actual, expected)) return true;
  if (actual.length !== expected.length) return false;

  return actual.every((segment, index) => {
    const want = expected[index];
    if (segment.trim().toLowerCase() === want.trim().toLowerCase()) {
      return true;
    }

    // Index 0 = company folder
    if (index === 0 && companyNumber?.trim()) {
      const prefix = `${companyNumber.trim().toLowerCase()} -`;
      return (
        segment.toLowerCase().startsWith(prefix) ||
        segment.toLowerCase() === companyNumber.trim().toLowerCase()
      );
    }

    // Index 2 = candidate folder when path is company/Candidates/candidate/type
    if (
      index === 2 &&
      workforceNumber?.trim() &&
      expected[1]?.toLowerCase() === "candidates"
    ) {
      const prefix = `${workforceNumber.trim().toLowerCase()} -`;
      return (
        segment.toLowerCase().startsWith(prefix) ||
        segment.toLowerCase() === workforceNumber.trim().toLowerCase()
      );
    }

    return false;
  });
}

/**
 * Read-only scan: reports files stored outside the folder expected from
 * company/candidate numbers + document type. Does not move or rename anything.
 */
export async function buildCustomerDocumentsMigrationReport(): Promise<DocumentMigrationReport> {
  const items = await getListItemsByKey("customerDocuments", { top: 5000 });

  const mismatches: DocumentMigrationMismatch[] = [];
  let matched = 0;
  let skipped = 0;
  let scannedFiles = 0;

  const companyCache = new Map<
    string,
    Awaited<ReturnType<typeof getCompanyById>>
  >();
  const workforceCache = new Map<
    string,
    Awaited<ReturnType<typeof getWorkforceById>>
  >();

  async function companyById(id: string) {
    if (!companyCache.has(id)) {
      companyCache.set(id, await getCompanyById(id));
    }
    return companyCache.get(id) ?? null;
  }

  async function workforceById(id: string) {
    if (!workforceCache.has(id)) {
      workforceCache.set(id, await getWorkforceById(id));
    }
    return workforceCache.get(id) ?? null;
  }

  for (const item of items) {
    if (isSharePointFolder(item.fields)) {
      skipped += 1;
      continue;
    }

    const fileName =
      asString(item.fields[documentFields.fileLeafRef]) ??
      asString(item.fields[documentFields.title]);
    if (!fileName) {
      skipped += 1;
      continue;
    }

    scannedFiles += 1;

    const companyId =
      asString(item.fields[documentFields.companyLookupId]) ?? null;
    const candidateId =
      asString(item.fields[documentFields.candidateLookupId]) ?? null;
    const companyLabel = asLookupOrString(item.fields[documentFields.company]);
    const candidateLabel = asLookupOrString(
      item.fields[documentFields.candidate],
    );
    const documentType = asNullableString(
      item.fields[documentFields.documentType],
    );
    const fileRef = asNullableString(item.fields[documentFields.fileRef]);
    const currentSegments = extractCustomerDocumentsFolderPath(fileRef);

    if (!companyId && !companyLabel) {
      mismatches.push({
        id: item.id,
        fileName,
        company: companyLabel,
        companyId,
        candidate: candidateLabel,
        candidateId,
        documentType,
        currentFolderPath: currentSegments.join("/"),
        expectedFolderPath: "",
        destinationFolder: "Other Documents",
        fileRef,
        reason: "Missing company metadata — cannot resolve expected folder.",
      });
      continue;
    }

    const company = companyId ? await companyById(companyId) : null;
    const companyName = company?.companyName ?? companyLabel;
    if (!companyName) {
      mismatches.push({
        id: item.id,
        fileName,
        company: companyLabel,
        companyId,
        candidate: candidateLabel,
        candidateId,
        documentType,
        currentFolderPath: currentSegments.join("/"),
        expectedFolderPath: "",
        destinationFolder: "Other Documents",
        fileRef,
        reason: "Company could not be resolved from metadata.",
      });
      continue;
    }

    const companyNumber = company?.companyNumber ?? null;
    let workforceNumber: string | null = null;
    let candidateName: string | null = candidateLabel;
    const hasCandidate = Boolean(candidateId);

    if (candidateId) {
      const workforce = await workforceById(candidateId);
      workforceNumber = workforce?.workforceNumber ?? null;
      candidateName = workforce?.candidateName ?? candidateLabel;
    }

    const { segments: expectedSegments, destinationFolder } =
      resolveDocumentFolderSegments({
        companyNumber,
        companyName,
        workforceNumber,
        candidateName,
        documentType,
        hasCandidate,
      });

    const expectedFolderPath = expectedSegments.join("/");
    const currentFolderPath = currentSegments.join("/");

    if (
      pathsMatchAllowingNumberStableNames(
        currentSegments,
        expectedSegments,
        companyNumber,
        workforceNumber,
      )
    ) {
      matched += 1;
      continue;
    }

    mismatches.push({
      id: item.id,
      fileName,
      company: companyName,
      companyId,
      candidate: candidateName,
      candidateId,
      documentType,
      currentFolderPath,
      expectedFolderPath,
      destinationFolder,
      fileRef,
      reason: currentSegments.length
        ? "Stored folder path does not match the expected structure for this document type."
        : "File has no resolvable folder path under Customer Documents.",
    });
  }

  return {
    generatedAt: new Date().toISOString(),
    scannedFiles,
    matched,
    mismatched: mismatches.length,
    skipped,
    mismatches,
    notes: [
      "This report does not move or rename files.",
      "Automatic migration is intentionally disabled until URL and integration impact is reviewed.",
      "Company and candidate folders are matched by stable number prefix when display names differ.",
    ],
  };
}

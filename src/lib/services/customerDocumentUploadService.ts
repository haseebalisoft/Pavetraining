import "server-only";

import { revalidateSharePointList } from "@/lib/cache/sharePointCache";
import { getGraphClient } from "@/lib/graph/graphClient";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { mapDocument } from "@/lib/services/adminCrudService";
import { getCompanyById } from "@/lib/services/companyService";
import {
  resolveDocumentUploadFolder,
  type DocumentDestinationFolder,
} from "@/lib/services/customerDocumentsFolderService";
import { ValidationError } from "@/lib/services/errorHandler";
import {
  getListItemByKey,
  toSharePointFields,
  updateListItemFieldsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import { getWorkforceById } from "@/lib/services/workforceService";
import {
  CUSTOMER_DOCUMENT_TYPES,
  type AdminDocumentRecord,
} from "@/types/adminDocuments";

const documentFields = getSharePointFields("customerDocuments");

/** 50 MB — within typical SharePoint Graph simple-upload limits. */
export const MAX_DOCUMENT_UPLOAD_BYTES = 50 * 1024 * 1024;

const ALLOWED_EXTENSIONS = new Set([
  "pdf",
  "png",
  "jpg",
  "jpeg",
  "gif",
  "webp",
  "tif",
  "tiff",
  "bmp",
  "doc",
  "docx",
  "xls",
  "xlsx",
  "ppt",
  "pptx",
  "txt",
  "csv",
  "rtf",
]);

export const DOCUMENT_TYPE_OPTIONS = CUSTOMER_DOCUMENT_TYPES;

export type UploadDocumentInput = {
  companyId: string;
  candidateId?: string | null;
  documentType: string;
  customerVisible?: boolean;
  /** When true, customers may be notified (subject to type/visibility rules). */
  notifyCustomer?: boolean;
  /** Bulk/import uploads should set true to suppress emails. */
  suppressNotifications?: boolean;
  fileName: string;
  contentType?: string | null;
  bytes: Uint8Array;
  actorEmail?: string | null;
};

export type UploadDocumentResult = {
  record: AdminDocumentRecord;
  folderPath: string;
  destinationFolder: DocumentDestinationFolder;
  notification?: import("@/types/notifications").DocumentNotificationResult | null;
};

function sanitizeFileName(fileName: string): string {
  const base = fileName.replace(/^.*[\\/]/, "").trim();
  const cleaned = base
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned || cleaned === "." || cleaned === "..") {
    throw new ValidationError("A valid file name is required.");
  }
  return cleaned.slice(0, 180);
}

function fileExtension(fileName: string): string {
  const idx = fileName.lastIndexOf(".");
  if (idx < 0) return "";
  return fileName.slice(idx + 1).toLowerCase();
}

export function validateDocumentUploadFile(input: {
  fileName: string;
  byteLength: number;
}): { fileName: string; extension: string } {
  const fileName = sanitizeFileName(input.fileName);
  if (input.byteLength <= 0) {
    throw new ValidationError("The uploaded file is empty.");
  }
  if (input.byteLength > MAX_DOCUMENT_UPLOAD_BYTES) {
    throw new ValidationError(
      `File exceeds the ${Math.floor(MAX_DOCUMENT_UPLOAD_BYTES / (1024 * 1024))} MB limit.`,
    );
  }
  const extension = fileExtension(fileName);
  if (!extension || !ALLOWED_EXTENSIONS.has(extension)) {
    throw new ValidationError(
      `File type ".${extension || "unknown"}" is not allowed. Allowed: ${Array.from(ALLOWED_EXTENSIONS).join(", ")}.`,
    );
  }
  return { fileName, extension };
}

function companiesMatch(
  candidateCompany: string | null | undefined,
  companyName: string,
): boolean {
  return (
    (candidateCompany ?? "").trim().toLowerCase() ===
    companyName.trim().toLowerCase()
  );
}

/**
 * Validates entities + file rules, ensures the correct Customer Documents
 * folder (by company/workforce number), uploads the binary, then writes
 * list metadata. Does not recreate the library.
 */
export async function uploadCustomerDocument(
  input: UploadDocumentInput,
): Promise<UploadDocumentResult> {
  const companyId = input.companyId?.trim();
  if (!companyId) {
    throw new ValidationError("Company is required.");
  }

  const documentType = input.documentType?.trim();
  if (!documentType) {
    throw new ValidationError("Document type is required.");
  }

  const company = await getCompanyById(companyId);
  if (!company) {
    throw new ValidationError("Company does not exist.");
  }

  const candidateId = input.candidateId?.trim() || null;
  const isCompanyLevel =
    !candidateId ||
    documentType.toLowerCase() === "company document" ||
    documentType.toLowerCase() === "company documents";

  let workforceNumber: string | null = null;
  let candidateName: string | null = null;

  if (!isCompanyLevel) {
    if (!candidateId) {
      throw new ValidationError(
        "Candidate is required for this document type.",
      );
    }
    const candidate = await getWorkforceById(candidateId);
    if (!candidate) {
      throw new ValidationError("Candidate does not exist.");
    }
    if (!companiesMatch(candidate.companyName, company.companyName)) {
      throw new ValidationError(
        "Candidate does not belong to the selected company.",
      );
    }
    workforceNumber = candidate.workforceNumber;
    candidateName = candidate.candidateName;
  }

  const { fileName } = validateDocumentUploadFile({
    fileName: input.fileName,
    byteLength: input.bytes.byteLength,
  });

  const folder = await resolveDocumentUploadFolder({
    companyNumber: company.companyNumber,
    companyName: company.companyName,
    workforceNumber,
    candidateName,
    documentType,
    hasCandidate: !isCompanyLevel,
  });

  if (!folder.folderId || folder.folderId === "root") {
    throw new ValidationError("Destination folder could not be resolved.");
  }

  const client = getGraphClient();
  const encodedName = encodeURIComponent(fileName);
  const uploaded = (await client
    .api(
      `/drives/${folder.driveId}/items/${folder.folderId}:/${encodedName}:/content`,
    )
    .header(
      "Content-Type",
      input.contentType?.trim() || "application/octet-stream",
    )
    .put(Buffer.from(input.bytes))) as {
    id?: string;
    name?: string;
  };

  if (!uploaded.id) {
    throw new Error("Upload succeeded without a drive item id.");
  }

  const listItem = (await client
    .api(`/drives/${folder.driveId}/items/${uploaded.id}/listItem`)
    .expand("fields")
    .get()) as { id?: string; fields?: SharePointFields };

  const listItemId = listItem.id ? String(listItem.id) : null;
  if (!listItemId) {
    throw new Error("Uploaded file has no SharePoint list item.");
  }

  const customerVisible = input.customerVisible !== false;
  const notifyCustomer =
    input.notifyCustomer === true ||
    (input.notifyCustomer !== false &&
      customerVisible &&
      !input.suppressNotifications);

  const metaPayload: SharePointFields = {
    ...toSharePointFields("customerDocuments", {
      title: fileName,
      documentType,
      customerVisible,
      notifyCustomer,
      notificationSent: false,
    }),
    [documentFields.companyLookupId]: Number(company.id),
  };

  if (!isCompanyLevel && candidateId) {
    metaPayload[documentFields.candidateLookupId] = Number(candidateId);
  }

  await updateListItemFieldsByKey("customerDocuments", listItemId, metaPayload);
  revalidateSharePointList("customerDocuments");

  const refreshed = await getListItemByKey("customerDocuments", listItemId);
  if (!refreshed) {
    throw new Error("Uploaded document could not be reloaded.");
  }

  const record = mapDocument(refreshed);
  if (!record) {
    throw new Error("Uploaded document could not be mapped.");
  }

  const { triggerDocumentNotificationSafe } = await import(
    "@/lib/services/documentNotificationService"
  );
  const notification = await triggerDocumentNotificationSafe(record, {
    suppressNotifications: Boolean(input.suppressNotifications),
    actorEmail: input.actorEmail,
  });

  const finalRecord =
    (await getListItemByKey("customerDocuments", listItemId).then((item) =>
      item ? mapDocument(item) : null,
    )) ?? record;

  return {
    record: finalRecord,
    folderPath: folder.path,
    destinationFolder: folder.destinationFolder,
    notification,
  };
}

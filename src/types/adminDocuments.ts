/**
 * Client-safe Customer Documents types/constants.
 * Keep out of adminCrudService so Client Components never import server-only.
 */

export type DocumentMetadataStatus =
  | "Complete"
  | "Missing Company"
  | "Missing Document Type"
  | "Hidden from Customer";

export interface AdminDocumentRecord {
  id: string;
  name: string;
  company: string | null;
  companyId: string | null;
  candidate: string | null;
  candidateId: string | null;
  documentType: string | null;
  customerVisible: boolean;
  /** No dedicated SharePoint column — true when the item is a file. */
  canDownload: boolean;
  notificationSent: boolean;
  notifyCustomer: boolean;
  modifiedDate: string | null;
  modifiedBy: string | null;
  metadataStatus: DocumentMetadataStatus;
  isFolder: boolean;
  fileRef: string | null;
  /** Relative folder segments under Customer Documents (parent of this item). */
  parentPath: string[];
  /** @deprecated Prefer modifiedDate — kept for older UI bindings. */
  uploadedDate: string | null;
  previewPath: string | null;
  downloadPath: string | null;
}

/** SharePoint Document Type choice values (Customer Documents). */
export const CUSTOMER_DOCUMENT_TYPES = [
  "Card Scan",
  "Certificate",
  "NVQ Document",
  "Training Brochure",
  "Other",
] as const;

export type CustomerDocumentType = (typeof CUSTOMER_DOCUMENT_TYPES)[number];

import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { uploadCustomerDocument } from "@/lib/services/customerDocumentUploadService";
import { NotFoundError, ValidationError } from "@/lib/services/errorHandler";
import {
  asBoolean,
  asString,
  buildFieldLookupIdEqualsFilter,
  extractLookupId,
  getListItemByKey,
  getListItemsByKey,
} from "@/lib/services/sharePointListService";
import {
  eventDocumentSubfolders,
  eventIdFromDocumentFolder,
} from "@/lib/training/eventDocumentFolders";

const documentFields = getSharePointFields("customerDocuments");
const eventFields = getSharePointFields("events");

/**
 * The Customer Documents "Document Type" choice used for event files. The
 * column is a fixed choice list; "Other" exists, and the Events/<id> folder
 * identifies the file as an event document.
 */
export const EVENT_DOCUMENT_TYPE = "Other";

export interface EventDocument {
  /** Customer Documents list item id — the existing view/download routes use it. */
  id: string;
  eventId: string;
  name: string;
  customerVisible: boolean;
  modified: string | null;
}

async function eventCompanyId(eventId: string): Promise<string> {
  const event = await getListItemByKey("events", eventId);
  if (!event) throw new NotFoundError("Event not found.");
  const companyId =
    extractLookupId(event.fields, eventFields.eventCompany) ||
    asString(event.fields[eventFields.eventCompanyLookupId]);
  if (!companyId) {
    throw new ValidationError(
      "This event has no company, so documents cannot be stored for it. Set the event's company first.",
    );
  }
  return String(companyId);
}

/**
 * Event documents for one company, grouped by event id. One list read per
 * company (filtered on the indexed Company lookup), then matched by folder.
 */
export async function listCompanyEventDocuments(
  companyId: string,
): Promise<Map<string, EventDocument[]>> {
  const id = Number(companyId);
  const items = Number.isFinite(id)
    ? await getListItemsByKey("customerDocuments", {
        filter: buildFieldLookupIdEqualsFilter(documentFields.companyLookupId, id),
        top: 5000,
        // Graph omits FileDirRef (the folder path that links a file to its
        // event) unless it is selected explicitly — found by the live test.
        selectFields: [
          documentFields.fileDirRef,
          documentFields.fileLeafRef,
          documentFields.title,
          documentFields.customerVisible,
          documentFields.fsObjType,
          documentFields.modified,
          documentFields.companyLookupId,
        ],
      })
    : [];
  const byEvent = new Map<string, EventDocument[]>();
  for (const item of items) {
    if (String(item.fields[documentFields.fsObjType] ?? "0") !== "0") continue;
    const eventId = eventIdFromDocumentFolder(
      asString(item.fields[documentFields.fileDirRef]),
    );
    if (!eventId) continue;
    const doc: EventDocument = {
      id: item.id,
      eventId,
      name:
        asString(item.fields[documentFields.fileLeafRef]) ??
        asString(item.fields[documentFields.title]) ??
        "document",
      customerVisible: asBoolean(item.fields[documentFields.customerVisible]),
      modified: asString(item.fields[documentFields.modified]) ?? null,
    };
    byEvent.set(eventId, [...(byEvent.get(eventId) ?? []), doc]);
  }
  for (const docs of byEvent.values()) {
    docs.sort((a, b) => a.name.localeCompare(b.name));
  }
  return byEvent;
}

/** Documents attached to one event (admin view: visible and hidden). */
export async function listEventDocuments(eventId: string): Promise<EventDocument[]> {
  const companyId = await eventCompanyId(eventId);
  return (await listCompanyEventDocuments(companyId)).get(String(eventId)) ?? [];
}

/**
 * Attach a file to an event through the existing document upload path
 * (Customer Documents, event's company folder, Company + visibility
 * metadata). No customer email is sent for event attachments.
 */
export async function uploadEventDocument(input: {
  eventId: string;
  fileName: string;
  contentType: string | null;
  bytes: Uint8Array;
  customerVisible: boolean;
  actorEmail: string | null;
}): Promise<EventDocument> {
  const companyId = await eventCompanyId(input.eventId);
  const result = await uploadCustomerDocument({
    companyId,
    documentType: EVENT_DOCUMENT_TYPE,
    companySubfolders: eventDocumentSubfolders(input.eventId),
    customerVisible: input.customerVisible,
    notifyCustomer: false,
    suppressNotifications: true,
    fileName: input.fileName,
    contentType: input.contentType,
    bytes: input.bytes,
    actorEmail: input.actorEmail,
  });
  return {
    id: result.record.id,
    eventId: String(input.eventId),
    name: result.record.name || input.fileName,
    customerVisible: input.customerVisible,
    modified: null,
  };
}

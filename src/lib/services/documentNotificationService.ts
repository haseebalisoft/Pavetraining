import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  mapDocument,
  updateAdminDocument,
} from "@/lib/services/adminCrudService";
import { getWorkforceById } from "@/lib/services/workforceService";
import {
  getPortalSettingsCached,
  isNotifiableDocumentType,
} from "@/lib/services/notificationConfig";
import { resolveNotificationRecipients } from "@/lib/services/notificationRecipientService";
import { sendAdminAlert, sendNotification } from "@/lib/services/notificationService";
import { documentUploadEmailTemplate } from "@/lib/services/notificationTemplateService";
import {
  getListItemByKey,
  listHasColumn,
  updateListItemFieldsByKey,
} from "@/lib/services/sharePointListService";
import type { AdminDocumentRecord } from "@/types/adminDocuments";
import type { DocumentNotificationResult } from "@/types/notifications";

const documentFields = getSharePointFields("customerDocuments");

export type DocumentNotificationOptions = {
  /** Bulk upload / import must pass true to suppress emails. */
  suppressNotifications?: boolean;
  actorEmail?: string | null;
  force?: boolean;
};

function skipResult(
  documentId: string,
  reason: string,
  notificationSent = false,
): DocumentNotificationResult {
  return {
    documentId,
    attempted: false,
    skipped: true,
    skipReason: reason,
    recipients: [],
    results: [],
    notificationSent,
  };
}

async function markDocumentNotificationState(
  documentId: string,
  state: {
    notificationSent: boolean;
    errorMessage?: string | null;
  },
): Promise<void> {
  const payload: Record<string, unknown> = {
    notificationSent: state.notificationSent,
  };

  try {
    await updateAdminDocument(documentId, payload);
  } catch (error) {
    console.error(
      "[document-notification] Failed to update NotificationSent",
      error,
    );
  }

  try {
    const extra: Record<string, unknown> = {};
    if (await listHasColumn("customerDocuments", "NotificationSentAt")) {
      if (state.notificationSent) {
        extra.NotificationSentAt = new Date().toISOString();
      }
    }
    if (await listHasColumn("customerDocuments", "NotificationError")) {
      extra.NotificationError = state.errorMessage?.slice(0, 250) ?? null;
    }
    if (Object.keys(extra).length > 0) {
      await updateListItemFieldsByKey("customerDocuments", documentId, extra);
    }
  } catch (error) {
    console.error(
      "[document-notification] Optional notification fields update failed",
      error,
    );
  }
}

export async function getAdminDocumentById(
  documentId: string,
): Promise<AdminDocumentRecord | null> {
  const item = await getListItemByKey("customerDocuments", documentId);
  if (!item) return null;
  return mapDocument(item);
}

/**
 * Sends customer document-upload notifications when eligibility rules pass.
 */
export async function notifyDocumentUpload(
  document: AdminDocumentRecord,
  options: DocumentNotificationOptions = {},
): Promise<DocumentNotificationResult> {
  if (options.suppressNotifications) {
    return skipResult(
      document.id,
      "SuppressNotifications=true — bulk/import upload suppressed emails.",
      document.notificationSent,
    );
  }

  const portal = await getPortalSettingsCached();

  if (!portal.enableCustomerNotifications || !portal.enableDocumentUploadNotifications) {
    return skipResult(
      document.id,
      "Document notifications disabled in Admin Settings.",
      document.notificationSent,
    );
  }

  if (document.isFolder) {
    return skipResult(document.id, "Folders are not notified.");
  }

  if (
    portal.requireCustomerVisibleBeforeNotification &&
    !document.customerVisible
  ) {
    return skipResult(document.id, "Document is hidden from customers.");
  }

  if (
    portal.requireNotifyCustomerBeforeNotification &&
    !document.notifyCustomer
  ) {
    return skipResult(document.id, "NotifyCustomer is not enabled.");
  }

  if (!isNotifiableDocumentType(document.documentType, portal)) {
    return skipResult(
      document.id,
      `DocumentType "${document.documentType ?? "—"}" is not enabled for notifications in Admin Settings.`,
    );
  }

  if (!document.companyId) {
    return skipResult(document.id, "Company is required before notifying.");
  }

  if (document.notificationSent && !options.force) {
    return skipResult(
      document.id,
      "NotificationSent is already true (duplicate prevented).",
      true,
    );
  }

  let department: string | null = null;
  if (document.candidateId) {
    try {
      const workforce = await getWorkforceById(document.candidateId);
      department = workforce?.department ?? null;
    } catch {
      department = null;
    }
  }

  const recipients = await resolveNotificationRecipients({
    companyId: document.companyId,
    audience: "document",
    candidateName: document.candidate,
    department,
  });

  if (recipients.length === 0) {
    await markDocumentNotificationState(document.id, {
      notificationSent: false,
      errorMessage: "No eligible recipients",
    });
    return {
      documentId: document.id,
      attempted: true,
      skipped: true,
      skipReason: "No eligible permission recipients for this company/scope.",
      recipients: [],
      results: [],
      notificationSent: false,
      errorMessage: "No eligible recipients",
    };
  }

  const template = documentUploadEmailTemplate({
    companyName: document.company ?? "your company",
    candidateName: document.candidate,
    documentType: document.documentType ?? "Document",
  });

  const results = [];
  for (const recipient of recipients) {
    results.push(
      await sendNotification({
        type: "document_upload",
        to: recipient.email,
        subject: template.subject,
        text: template.text,
        html: template.html,
        companyName: document.company,
        itemId: document.id,
        actorEmail: options.actorEmail,
        dedupeKey: `document:${document.id}:${recipient.email}`,
        detail: `DocumentType=${document.documentType}; Candidate=${document.candidate ?? "—"}`,
      }),
    );
  }

  const anyFailed = results.some((r) => r.status === "failed");
  const anySentOrLogged = results.some(
    (r) =>
      r.status === "sent" ||
      r.status === "not_configured" ||
      r.status === "queued",
  );
  const errorMessage = anyFailed
    ? results
        .filter((r) => r.status === "failed")
        .map((r) => r.errorMessage)
        .filter(Boolean)
        .join("; ")
    : null;

  // Mark sent when we successfully delivered OR safely logged not_configured
  // so we do not spam retries. Force can resend.
  if (anySentOrLogged) {
    await markDocumentNotificationState(document.id, {
      notificationSent: true,
      errorMessage,
    });
  } else if (anyFailed) {
    await markDocumentNotificationState(document.id, {
      notificationSent: false,
      errorMessage,
    });
    await sendAdminAlert({
      title: "Failed document notification",
      detail: `Document #${document.id} (${document.name}): ${errorMessage}`,
      itemId: document.id,
      actorEmail: options.actorEmail,
    });
  }

  return {
    documentId: document.id,
    attempted: true,
    skipped: false,
    recipients: recipients.map((r) => r.email),
    results,
    notificationSent: anySentOrLogged,
    errorMessage,
  };
}

export async function notifyDocumentById(
  documentId: string,
  options: DocumentNotificationOptions = {},
): Promise<DocumentNotificationResult> {
  const document = await getAdminDocumentById(documentId);
  if (!document) {
    return skipResult(documentId, "Document not found.");
  }
  return notifyDocumentUpload(document, options);
}

/**
 * After upload/metadata update — best effort, never throws to callers.
 */
export async function triggerDocumentNotificationSafe(
  document: AdminDocumentRecord,
  options: DocumentNotificationOptions = {},
): Promise<DocumentNotificationResult | null> {
  try {
    return await notifyDocumentUpload(document, options);
  } catch (error) {
    console.error("[document-notification] Unexpected failure", error);
    try {
      await sendAdminAlert({
        title: "Failed document notification",
        detail:
          error instanceof Error
            ? error.message
            : "Unexpected document notification error",
        itemId: document.id,
        actorEmail: options.actorEmail,
      });
    } catch {
      // ignore
    }
    return null;
  }
}

/** Expose field internal names for diagnostics. */
export function documentNotificationFieldNames() {
  return {
    notifyCustomer: documentFields.notifyCustomer,
    notificationSent: documentFields.notificationSent,
    customerVisible: documentFields.customerVisible,
    documentType: documentFields.documentType,
  };
}

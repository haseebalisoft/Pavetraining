import "server-only";

import { getGraphClient } from "@/lib/graph/graphClient";
import { getNotificationSettings, getPortalSettingsCached } from "@/lib/services/notificationConfig";
import { writeNotificationLog } from "@/lib/services/notificationLogService";
import {
  adminAlertEmailTemplate,
  loadPaveLogoAttachment,
  portalInviteEmailTemplate,
  testEmailTemplate,
} from "@/lib/services/notificationTemplateService";
import { resolveAdminAlertRecipients } from "@/lib/services/notificationRecipientService";
import type {
  NotificationDeliveryStatus,
  NotificationSendResult,
  NotificationType,
} from "@/types/notifications";

export interface SendNotificationInput {
  type: NotificationType;
  to: string;
  subject: string;
  text: string;
  html?: string;
  companyName?: string | null;
  itemId?: string | null;
  dedupeKey?: string | null;
  actorEmail?: string | null;
  detail?: string | null;
  /** Optional file attachments (e.g. .ics calendar invite or inline logo). */
  attachments?: Array<{
    filename: string;
    contentType: string;
    /** UTF-8 text or already-encoded base64 content. */
    content: string;
    encoding?: "utf8" | "base64";
    /** When set with isInline, referenced from HTML as cid:{contentId}. */
    contentId?: string;
    isInline?: boolean;
  }>;
  /** Optional From display name (mailbox address still uses NOTIFICATION_FROM_EMAIL). */
  fromName?: string | null;
}

async function sendViaGraph(input: {
  from: string;
  fromName?: string | null;
  to: string;
  subject: string;
  text: string;
  html?: string;
  attachments?: SendNotificationInput["attachments"];
}): Promise<{ ok: true; messageId?: string } | { ok: false; error: string }> {
  try {
    const client = getGraphClient();
    const attachments =
      input.attachments?.map((file) => {
        const contentBytes =
          file.encoding === "base64"
            ? file.content
            : Buffer.from(file.content, "utf8").toString("base64");
        return {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: file.filename,
          contentType: file.contentType,
          contentBytes,
          ...(file.contentId
            ? {
                contentId: file.contentId,
                isInline: file.isInline !== false,
              }
            : {}),
        };
      }) ?? [];

    await client.api(`/users/${encodeURIComponent(input.from)}/sendMail`).post({
      message: {
        subject: input.subject,
        body: {
          contentType: input.html ? "HTML" : "Text",
          content: input.html ?? input.text,
        },
        from: {
          emailAddress: {
            address: input.from,
            name: input.fromName?.trim() || "PAVE Training Portal",
          },
        },
        toRecipients: [
          {
            emailAddress: {
              address: input.to,
            },
          },
        ],
        ...(attachments.length > 0 ? { attachments } : {}),
      },
      saveToSentItems: false,
    });
    return { ok: true };
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Graph sendMail failed.";
    return { ok: false, error: message };
  }
}

/**
 * Sends a notification email when configured; otherwise logs a safe fallback.
 * Never throws for delivery failures — returns status instead.
 */
export async function sendNotification(
  input: SendNotificationInput,
): Promise<NotificationSendResult> {
  const settings = await getNotificationSettings();
  const portal = await getPortalSettingsCached();
  const to = input.to.trim().toLowerCase();

  const isAdminAlert = input.type === "admin_alert";
  const isPortalInvite = input.type === "portal_invite";
  const isLoginOtp = input.type === "login_otp";
  const isBookingConfirmed = input.type === "booking_confirmed";
  const isExpiry =
    input.type === "expiry_3m" ||
    input.type === "expiry_6m" ||
    input.type === "expiry_expired";

  if (isAdminAlert && !portal.enableAdminAlerts) {
    await writeNotificationLog({
      type: input.type,
      status: "skipped",
      recipientEmail: to,
      companyName: input.companyName,
      subject: input.subject,
      dedupeKey: input.dedupeKey,
      itemId: input.itemId,
      errorMessage: "Admin alerts disabled by settings.",
      detail: input.detail,
      actorEmail: input.actorEmail,
    });
    return {
      status: "skipped",
      recipientEmail: to,
      subject: input.subject,
      errorMessage: "Admin alerts disabled by settings.",
      logged: true,
    };
  }

  // Portal invites + login OTP must send when mail is configured.
  // Booking confirmations only need the customer master switch (not document-upload).
  if (
    !isAdminAlert &&
    !isPortalInvite &&
    !isLoginOtp &&
    (!settings.notificationsEnabled || !portal.enableCustomerNotifications)
  ) {
    await writeNotificationLog({
      type: input.type,
      status: "skipped",
      recipientEmail: to,
      companyName: input.companyName,
      subject: input.subject,
      dedupeKey: input.dedupeKey,
      itemId: input.itemId,
      errorMessage: "Customer notifications disabled by settings.",
      detail: input.detail,
      actorEmail: input.actorEmail,
    });
    return {
      status: "skipped",
      recipientEmail: to,
      subject: input.subject,
      errorMessage: "Notifications disabled by settings.",
      logged: true,
    };
  }

  if (
    input.type === "document_upload" &&
    !portal.enableDocumentUploadNotifications
  ) {
    await writeNotificationLog({
      type: input.type,
      status: "skipped",
      recipientEmail: to,
      companyName: input.companyName,
      subject: input.subject,
      dedupeKey: input.dedupeKey,
      itemId: input.itemId,
      errorMessage: "Document upload notifications disabled by settings.",
      detail: input.detail,
      actorEmail: input.actorEmail,
    });
    return {
      status: "skipped",
      recipientEmail: to,
      subject: input.subject,
      errorMessage: "Document upload notifications disabled by settings.",
      logged: true,
    };
  }

  if (isBookingConfirmed && !portal.enableCustomerNotifications) {
    await writeNotificationLog({
      type: input.type,
      status: "skipped",
      recipientEmail: to,
      companyName: input.companyName,
      subject: input.subject,
      dedupeKey: input.dedupeKey,
      itemId: input.itemId,
      errorMessage: "Customer notifications disabled by settings.",
      detail: input.detail,
      actorEmail: input.actorEmail,
    });
    return {
      status: "skipped",
      recipientEmail: to,
      subject: input.subject,
      errorMessage: "Customer notifications disabled by settings.",
      logged: true,
    };
  }

  if (isExpiry && !portal.enableExpiryReminders) {
    await writeNotificationLog({
      type: input.type,
      status: "skipped",
      recipientEmail: to,
      companyName: input.companyName,
      subject: input.subject,
      dedupeKey: input.dedupeKey,
      itemId: input.itemId,
      errorMessage: "Expiry reminders disabled by settings.",
      detail: input.detail,
      actorEmail: input.actorEmail,
    });
    return {
      status: "skipped",
      recipientEmail: to,
      subject: input.subject,
      errorMessage: "Expiry reminders disabled by settings.",
      logged: true,
    };
  }

  if (!settings.emailConfigured || !settings.fromEmail) {
    await writeNotificationLog({
      type: input.type,
      status: "not_configured",
      recipientEmail: to,
      companyName: input.companyName,
      subject: input.subject,
      dedupeKey: input.dedupeKey,
      itemId: input.itemId,
      errorMessage: "Not sent - email not configured",
      detail: input.detail ?? input.text.slice(0, 500),
      actorEmail: input.actorEmail,
    });
    return {
      status: "not_configured",
      recipientEmail: to,
      subject: input.subject,
      errorMessage: "Not sent - email not configured",
      logged: true,
    };
  }

  const delivery = await sendViaGraph({
    from: settings.fromEmail,
    fromName: input.fromName,
    to,
    subject: input.subject,
    text: input.text,
    html: input.html,
    attachments: input.attachments,
  });

  if (!delivery.ok) {
    await writeNotificationLog({
      type: input.type,
      status: "failed",
      recipientEmail: to,
      companyName: input.companyName,
      subject: input.subject,
      dedupeKey: input.dedupeKey,
      itemId: input.itemId,
      errorMessage: delivery.error,
      detail: input.detail,
      actorEmail: input.actorEmail,
    });
    const { logNotification } = await import("@/lib/services/auditLogService");
    await logNotification({
      userEmail: input.actorEmail,
      recipientEmail: to,
      success: false,
      company: input.companyName,
      entityId: input.itemId,
      entityName: input.subject,
      errorMessage: delivery.error,
      metadata: { type: input.type },
    });
    return {
      status: "failed",
      recipientEmail: to,
      subject: input.subject,
      errorMessage: delivery.error,
      logged: true,
    };
  }

  await writeNotificationLog({
    type: input.type,
    status: "sent",
    recipientEmail: to,
    companyName: input.companyName,
    subject: input.subject,
    dedupeKey: input.dedupeKey,
    itemId: input.itemId,
    detail: input.detail,
    actorEmail: input.actorEmail,
  });
  const { logNotification } = await import("@/lib/services/auditLogService");
  await logNotification({
    userEmail: input.actorEmail,
    recipientEmail: to,
    success: true,
    company: input.companyName,
    entityId: input.itemId,
    entityName: input.subject,
    metadata: { type: input.type },
  });

  return {
    status: "sent",
    recipientEmail: to,
    subject: input.subject,
    messageId: null,
    logged: true,
  };
}

export async function sendTestNotification(input: {
  to: string;
  actorEmail?: string | null;
}): Promise<NotificationSendResult> {
  const template = testEmailTemplate();
  return sendNotification({
    type: "test",
    to: input.to,
    subject: template.subject,
    text: template.text,
    html: template.html,
    actorEmail: input.actorEmail,
    detail: "Manual admin test email",
  });
}

export async function sendAdminAlert(input: {
  title: string;
  detail: string;
  itemId?: string | null;
  actorEmail?: string | null;
}): Promise<NotificationSendResult[]> {
  const template = adminAlertEmailTemplate({
    title: input.title,
    detail: input.detail,
  });
  const recipients = await resolveAdminAlertRecipients();
  if (recipients.length === 0) {
    await writeNotificationLog({
      type: "admin_alert",
      status: "skipped",
      subject: template.subject,
      itemId: input.itemId,
      errorMessage: "No admin alert recipients configured.",
      detail: input.detail,
      actorEmail: input.actorEmail,
    });
    return [];
  }

  const results: NotificationSendResult[] = [];
  for (const email of recipients) {
    results.push(
      await sendNotification({
        type: "admin_alert",
        to: email,
        subject: template.subject,
        text: template.text,
        html: template.html,
        itemId: input.itemId,
        actorEmail: input.actorEmail,
        detail: input.detail,
        dedupeKey: `admin-alert:${input.title}:${input.itemId ?? "none"}:${email}:${new Date().toISOString().slice(0, 10)}`,
      }),
    );
  }
  return results;
}

export async function sendPortalInviteNotification(input: {
  to: string;
  displayName?: string | null;
  companyName?: string | null;
  roleLabel?: string | null;
  itemId?: string | null;
  actorEmail?: string | null;
}): Promise<NotificationSendResult> {
  const template = portalInviteEmailTemplate({
    displayName: input.displayName,
    companyName: input.companyName,
    roleLabel: input.roleLabel,
  });
  const logo = await loadPaveLogoAttachment();
  return sendNotification({
    type: "portal_invite",
    to: input.to,
    subject: template.subject,
    text: template.text,
    html: template.html,
    companyName: input.companyName,
    itemId: input.itemId,
    actorEmail: input.actorEmail,
    detail: "Permissions list invite",
    fromName: "PAVE Training",
    dedupeKey: `portal-invite:${input.to.toLowerCase()}:${input.itemId ?? "new"}`,
    attachments: logo ? [logo] : undefined,
  });
}

export function summarizeStatuses(
  results: NotificationSendResult[],
): Record<NotificationDeliveryStatus, number> {
  const counts: Record<NotificationDeliveryStatus, number> = {
    sent: 0,
    queued: 0,
    skipped: 0,
    failed: 0,
    not_configured: 0,
  };
  for (const result of results) {
    counts[result.status] += 1;
  }
  return counts;
}

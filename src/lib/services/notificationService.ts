import "server-only";

import { getGraphClient } from "@/lib/graph/graphClient";
import { getNotificationSettings, getPortalSettingsCached } from "@/lib/services/notificationConfig";
import { writeNotificationLog } from "@/lib/services/notificationLogService";
import {
  adminAlertEmailTemplate,
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
}

async function sendViaGraph(input: {
  from: string;
  to: string;
  subject: string;
  text: string;
  html?: string;
}): Promise<{ ok: true; messageId?: string } | { ok: false; error: string }> {
  try {
    const client = getGraphClient();
    await client.api(`/users/${encodeURIComponent(input.from)}/sendMail`).post({
      message: {
        subject: input.subject,
        body: {
          contentType: input.html ? "HTML" : "Text",
          content: input.html ?? input.text,
        },
        toRecipients: [
          {
            emailAddress: {
              address: input.to,
            },
          },
        ],
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

  if (
    !isAdminAlert &&
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
    to,
    subject: input.subject,
    text: input.text,
    html: input.html,
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

import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { hasRecentNotificationDedupe } from "@/lib/services/notificationLogService";
import { sendNotification } from "@/lib/services/notificationService";
import { emailLogoHtml } from "@/lib/services/notificationTemplateService";
import {
  asString,
  extractLookupId,
  getListItemByKey,
} from "@/lib/services/sharePointListService";
import { getWorkforceIdByCandidateName } from "@/lib/services/workforceService";
import { formatDate } from "@/lib/utils/formatDate";
import type { NotificationSendResult } from "@/types/notifications";

/**
 * Emails the assigned Training Manager whenever a training record is added,
 * updated, or confirmed (Pass/Fail outcome saved).
 *
 * Covers NPORS, EUSR, Streetworks / NRSWA, In-House certificates and NVQ.
 * The recipient is resolved from the candidate's Workforce row —
 * `Trainingmanager` Lookup → Permissions row → user email — so the message
 * always goes to the Training Manager actually responsible for that candidate,
 * not just any TM at the company.
 *
 * This service never throws — every failure path is logged and returned as a
 * skipped / failed result so the underlying save flow is never blocked.
 */

type Register =
  | "nporsRegister"
  | "eusrRegister"
  | "nrswaRegister"
  | "inHouseCertificates"
  | "nvqRegister";

type ChangeAction = "added" | "updated" | "confirmed";

/** Fields the notification pulls from the just-saved record. */
export interface TrainingRecordNotificationRecord {
  id: string;
  candidateName: string;
  companyName: string | null;
  /** SharePoint Workforce Lookup id — required to resolve the assigned TM. */
  candidateLookupId?: string | null;
  /** Alias used by some register mappers. */
  workforceId?: string | null;
  trainingDate?: string | null;
  /** NVQ uses registration / induction instead of a training date. */
  dateRegistered?: string | null;
  inductionDate?: string | null;
  expiry?: string | null;
  /** Register-specific category / course fields — first non-empty wins. */
  nporsCategory?: string | null;
  eusrCategory?: string | null;
  streetworksCategory?: string | null;
  certificateCategory?: string | null;
  courseCategory?: string | null;
  course?: string | null;
  /** NVQ record shape. */
  nvqTitle?: string | null;
  boltOn?: string | null;
  /** Optional outcome text used to detect "confirmed" (Pass / Fail). */
  trainingOutcome?: string | null;
  /** NVQ status — Completed is treated as confirmed. */
  status?: string | null;
}

export interface NotifyTrainingRecordResult {
  attempted: boolean;
  skipped: boolean;
  skipReason?: string;
  recipient?: string;
  result?: NotificationSendResult;
}

const REGISTER_LABELS: Record<Register, string> = {
  nporsRegister: "NPORS",
  eusrRegister: "EUSR",
  nrswaRegister: "Streetworks",
  inHouseCertificates: "In-House",
  nvqRegister: "NVQ",
};

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function isConfirmedOutcome(outcome: string | null | undefined): boolean {
  const value = outcome?.trim().toLowerCase();
  return value === "pass" || value === "fail" || value === "confirmed";
}

function resolveAction(
  requested: "added" | "updated",
  record: TrainingRecordNotificationRecord,
): ChangeAction {
  if (requested !== "updated") return requested;
  if (isConfirmedOutcome(record.trainingOutcome)) return "confirmed";
  if (record.status?.trim().toLowerCase() === "completed") return "confirmed";
  return requested;
}

function resolveTrainingDate(record: TrainingRecordNotificationRecord): string | null {
  return (
    record.trainingDate?.trim() ||
    record.dateRegistered?.trim() ||
    record.inductionDate?.trim() ||
    null
  );
}

async function resolveWorkforceId(
  record: TrainingRecordNotificationRecord,
): Promise<string | null> {
  const direct =
    record.candidateLookupId?.trim() || record.workforceId?.trim() || "";
  if (direct) return direct;

  const company = record.companyName?.trim();
  const name = record.candidateName?.trim();
  if (!company || !name) return null;

  const map = await getWorkforceIdByCandidateName(company);
  return map.get(name.toLowerCase()) ?? null;
}

function resolveCategory(
  register: Register,
  record: TrainingRecordNotificationRecord,
): string {
  const candidates = [
    record.nporsCategory,
    record.eusrCategory,
    record.streetworksCategory,
    record.certificateCategory,
    record.courseCategory,
    record.course,
    record.nvqTitle,
    record.boltOn,
  ];
  for (const value of candidates) {
    const trimmed = value?.trim();
    if (trimmed) return trimmed;
  }
  return REGISTER_LABELS[register];
}

/**
 * Reads the candidate's Workforce row and returns the assigned Training
 * Manager's Permissions row id, or null when there is no TM assignment.
 */
async function resolveTrainingManagerLookupId(
  candidateLookupId: string,
): Promise<string | null> {
  const workforceFields = getSharePointFields("workforce");
  const workforce = await getListItemByKey("workforce", candidateLookupId);
  if (!workforce) return null;
  const tmLookupId = extractLookupId(
    workforce.fields,
    workforceFields.trainingManager,
  );
  return tmLookupId || null;
}

interface TrainingManagerContact {
  email: string;
  displayName: string | null;
}

/**
 * Reads the Permissions row for the assigned TM and returns email + name.
 * Skips inactive TMs so we never email someone whose portal access is off.
 */
async function loadTrainingManagerContact(
  permissionId: string,
): Promise<TrainingManagerContact | null> {
  const permissionFields = getSharePointFields("permissions");
  const permission = await getListItemByKey("permissions", permissionId);
  if (!permission) return null;

  const status = asString(permission.fields[permissionFields.status])?.trim();
  if (status && status.toLowerCase() !== "active") return null;

  const email = asString(permission.fields[permissionFields.userEmail])
    ?.trim()
    .toLowerCase();
  if (!email) return null;

  const displayName =
    asString(permission.fields[permissionFields.name])?.trim() || null;

  return { email, displayName };
}

/**
 * Compose the email body. Keeps the exact contract the client specified:
 *   Candidate name / Training category / Training date / Expiry date /
 *   Company / Updated by
 */
function buildEmail(input: {
  action: ChangeAction;
  register: Register;
  record: TrainingRecordNotificationRecord;
  category: string;
  actor: string;
  tmName: string | null;
}): { subject: string; text: string; html: string } {
  const registerLabel = REGISTER_LABELS[input.register];
  const actionLabel =
    input.action === "added"
      ? "added"
      : input.action === "confirmed"
        ? "confirmed"
        : "updated";
  const candidate = input.record.candidateName?.trim() || "—";
  const company = input.record.companyName?.trim() || "—";
  const trainingDate = formatDate(resolveTrainingDate(input.record));
  const expiry =
    input.register === "nvqRegister"
      ? "N/A (NVQ)"
      : formatDate(input.record.expiry);
  const outcome = input.record.trainingOutcome?.trim();
  const greeting = input.tmName ? `Hi ${input.tmName},` : "Hi,";

  const subject = `Training ${actionLabel}: ${candidate} — ${registerLabel} · ${input.category}`;

  const text = [
    greeting,
    "",
    `A ${registerLabel} training record has been ${actionLabel} in the PAVE Training Portal.`,
    "",
    `Candidate name: ${candidate}`,
    `Training category: ${input.category}`,
    `Training date: ${trainingDate}`,
    `Expiry date: ${expiry}`,
    `Company: ${company}`,
    outcome ? `Outcome: ${outcome}` : null,
    `Updated by: ${input.actor}`,
    "",
    "Sign in to the PAVE Training Portal to review the record and any candidate documents.",
  ]
    .filter((line): line is string => line !== null)
    .join("\n");

  const html = `${emailLogoHtml()}<p>${escapeHtml(greeting)}</p>
<p>A <strong>${escapeHtml(registerLabel)}</strong> training record has been <strong>${escapeHtml(actionLabel)}</strong> in the PAVE Training Portal.</p>
<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
  <tr><td style="padding:2px 12px 2px 0"><strong>Candidate name</strong></td><td style="padding:2px 0">${escapeHtml(candidate)}</td></tr>
  <tr><td style="padding:2px 12px 2px 0"><strong>Training category</strong></td><td style="padding:2px 0">${escapeHtml(input.category)}</td></tr>
  <tr><td style="padding:2px 12px 2px 0"><strong>Training date</strong></td><td style="padding:2px 0">${escapeHtml(trainingDate)}</td></tr>
  <tr><td style="padding:2px 12px 2px 0"><strong>Expiry date</strong></td><td style="padding:2px 0">${escapeHtml(expiry)}</td></tr>
  <tr><td style="padding:2px 12px 2px 0"><strong>Company</strong></td><td style="padding:2px 0">${escapeHtml(company)}</td></tr>
  ${outcome ? `<tr><td style="padding:2px 12px 2px 0"><strong>Outcome</strong></td><td style="padding:2px 0">${escapeHtml(outcome)}</td></tr>` : ""}
  <tr><td style="padding:2px 12px 2px 0"><strong>Updated by</strong></td><td style="padding:2px 0">${escapeHtml(input.actor)}</td></tr>
</table>
<p>Sign in to the PAVE Training Portal to review the record and any candidate documents.</p>`;

  return { subject, text, html };
}

/**
 * Fire-and-forget notification for a single register save. Best-effort:
 * catches any error so the underlying training-record save is never blocked.
 */
export async function notifyTrainingRecordChange(input: {
  register: Register;
  action: "added" | "updated";
  record: TrainingRecordNotificationRecord;
  actorEmail: string | null | undefined;
}): Promise<NotifyTrainingRecordResult> {
  try {
    const action = resolveAction(input.action, input.record);
    const workforceId = await resolveWorkforceId(input.record);
    if (!workforceId) {
      return {
        attempted: false,
        skipped: true,
        skipReason:
          "Candidate is not linked to a Workforce row, so no Training Manager can be resolved.",
      };
    }

    const tmPermissionId = await resolveTrainingManagerLookupId(workforceId);
    if (!tmPermissionId) {
      return {
        attempted: false,
        skipped: true,
        skipReason:
          "Candidate has no Training Manager assigned on the Workforce row.",
      };
    }

    const contact = await loadTrainingManagerContact(tmPermissionId);
    if (!contact) {
      return {
        attempted: false,
        skipped: true,
        skipReason:
          "Assigned Training Manager could not be resolved (missing Permissions row, inactive, or no email).",
      };
    }

    const category = resolveCategory(input.register, input.record);
    const actor = input.actorEmail?.trim() || "PAVE Training";

    // Dedupe per action so "added" then "confirmed" both send. The same
    // action on the same record within 10 minutes is collapsed.
    const window = Math.floor(Date.now() / (10 * 60_000));
    const dedupeKey = `training-notify:${input.register}:${input.record.id}:${action}:${contact.email}:${window}`;
    if (await hasRecentNotificationDedupe(dedupeKey, 1)) {
      return {
        attempted: false,
        skipped: true,
        skipReason:
          "Recent notification already sent for this record and action within the 10-minute dedupe window.",
        recipient: contact.email,
      };
    }

    const template = buildEmail({
      action,
      register: input.register,
      record: input.record,
      category,
      actor,
      tmName: contact.displayName,
    });

    const result = await sendNotification({
      type: "training_record_change",
      to: contact.email,
      subject: template.subject,
      text: template.text,
      html: template.html,
      companyName: input.record.companyName ?? null,
      itemId: input.record.id,
      dedupeKey,
      actorEmail: input.actorEmail,
      detail: `${REGISTER_LABELS[input.register]} training ${action}`,
      fromName: "PAVE Training",
    });

    return {
      attempted: true,
      skipped: result.status === "skipped",
      skipReason:
        result.status === "skipped" ? (result.errorMessage ?? undefined) : undefined,
      recipient: contact.email,
      result,
    };
  } catch (error) {
    // Never throw — the caller must not fail the save because of a mail error.
    const message =
      error instanceof Error ? error.message : "Unknown notification error.";
    console.error(
      `[notifyTrainingRecordChange] ${input.register} ${input.action} #${input.record.id}: ${message}`,
    );
    return {
      attempted: true,
      skipped: true,
      skipReason: `Notification failed: ${message}`,
    };
  }
}

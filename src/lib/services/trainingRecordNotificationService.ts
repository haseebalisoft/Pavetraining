import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  hasRecentNotificationDedupe,
  writeNotificationLog,
} from "@/lib/services/notificationLogService";
import { mapPermissionItem } from "@/lib/services/permissionService";
import { sendNotification } from "@/lib/services/notificationService";
import {
  loadEmailBrandAttachments,
  trainingRecordChangeEmailTemplate,
} from "@/lib/services/notificationTemplateService";
import {
  formatNporsCategoryLabel,
  splitNporsCategories,
} from "@/lib/training/nporsCategoryOptions";
import {
  extractLookupId,
  getListItemByKey,
} from "@/lib/services/sharePointListService";
import { getWorkforceIdByCandidateName } from "@/lib/services/workforceService";
import { formatDate } from "@/lib/utils/formatDate";
import {
  checkTrainingManagerContact,
  resolveAction,
  summarizeTrainingManagerEmail,
  type ChangeAction,
  type NotifyTrainingRecordResult,
  type TrainingRecordNotificationRecord,
} from "@/lib/services/trainingManagerNotificationRules";

export {
  checkTrainingManagerContact,
  resolveAction,
  summarizeTrainingManagerEmail,
};
export type {
  NotifyTrainingRecordResult,
  TrainingManagerEmailSummary,
  TrainingRecordNotificationRecord,
} from "@/lib/services/trainingManagerNotificationRules";
import type { PermissionProfile } from "@/types/models";
import type { NotificationCategory } from "@/types/notifications";

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



const REGISTER_CATEGORIES: Record<Register, NotificationCategory> = {
  nporsRegister: "training_npors",
  eusrRegister: "training_eusr",
  nrswaRegister: "training_streetworks",
  inHouseCertificates: "training_in_house",
  nvqRegister: "training_nvq",
};

const REGISTER_LABELS: Record<Register, string> = {
  nporsRegister: "NPORS",
  eusrRegister: "EUSR",
  nrswaRegister: "Streetworks",
  inHouseCertificates: "In-House",
  nvqRegister: "NVQ",
};


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
 * Reads the candidate's Workforce row: the assigned Training Manager's
 * Permissions row id (null when none) and the candidate's company id.
 */
async function resolveAssignment(
  candidateLookupId: string,
): Promise<{ tmPermissionId: string | null; companyId: string | null } | null> {
  const workforceFields = getSharePointFields("workforce");
  const workforce = await getListItemByKey("workforce", candidateLookupId);
  if (!workforce) return null;
  return {
    tmPermissionId:
      extractLookupId(workforce.fields, workforceFields.trainingManager) ||
      null,
    companyId:
      extractLookupId(workforce.fields, workforceFields.companyName) || null,
  };
}


async function loadTrainingManagerProfile(
  permissionId: string,
): Promise<PermissionProfile | null> {
  const permission = await getListItemByKey("permissions", permissionId);
  if (!permission) return null;
  return mapPermissionItem(permission.id, permission.fields);
}

/**
 * Compose the email body. Keeps the exact contract the client specified:
 *   Candidate name / Training category / Training date / Expiry date /
 *   Company / Updated by — plus a portal link for external customers.
 */
function buildEmail(input: {
  action: ChangeAction;
  register: Register;
  record: TrainingRecordNotificationRecord;
  category: string;
  actor: string;
  tmName: string | null;
  includeCompanyLogo?: boolean;
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

  // NPORS: every category on the record as "N001 – Ind FLT", so the N
  // number never appears without its category name.
  const nporsLabels =
    input.register === "nporsRegister"
      ? splitNporsCategories(input.record.nporsCategory)
          .map(formatNporsCategoryLabel)
          .filter(Boolean)
      : [];
  const categoryName = input.category;
  const nNumber =
    nporsLabels.length > 0
      ? nporsLabels.join(", ")
      : input.register === "nporsRegister"
        ? input.record.nporsNumber?.trim() || null
        : null;

  return trainingRecordChangeEmailTemplate({
    greeting,
    registerLabel,
    actionLabel,
    candidate,
    category: input.category,
    categoryName,
    nNumber,
    trainingDate,
    expiry,
    company,
    outcome,
    actor: input.actor,
    includeCompanyLogo: input.includeCompanyLogo,
  });
}

/**
 * Notification for a single register save. Best-effort: never throws, so the
 * training-record save is never blocked. Every path that does not email the
 * Training Manager is written to the notification log with its reason.
 */
export async function notifyTrainingRecordChange(input: {
  register: Register;
  action: "added" | "updated";
  record: TrainingRecordNotificationRecord;
  actorEmail: string | null | undefined;
  /**
   * Pass/Fail result before this edit (PATCH only). Lets an unchanged result
   * go out as "updated" instead of a repeated result email.
   */
  previousOutcome?: string | null;
}): Promise<NotifyTrainingRecordResult> {
  const registerLabel = REGISTER_LABELS[input.register];
  const skip = async (skipReason: string): Promise<NotifyTrainingRecordResult> => {
    await writeNotificationLog({
      type: "training_record_change",
      status: "skipped",
      recipientEmail: null,
      companyName: input.record.companyName ?? null,
      subject: `${registerLabel} training ${input.action}: ${input.record.candidateName}`,
      itemId: input.record.id,
      errorMessage: skipReason,
      detail: `${registerLabel} training ${input.action}`,
      actorEmail: input.actorEmail,
    });
    return { attempted: false, skipped: true, skipReason };
  };

  try {
    const action = resolveAction(
      input.action,
      input.record,
      input.previousOutcome,
    );
    const workforceId = await resolveWorkforceId(input.record);
    if (!workforceId) {
      return await skip(
        "Candidate is not linked to a Workforce row, so no Training Manager can be resolved.",
      );
    }

    const assignment = await resolveAssignment(workforceId);
    if (!assignment) {
      return await skip(`Workforce row #${workforceId} was not found.`);
    }
    if (!assignment.tmPermissionId) {
      return await skip(
        "Candidate has no Training Manager assigned on the Workforce row.",
      );
    }

    const checked = checkTrainingManagerContact(
      await loadTrainingManagerProfile(assignment.tmPermissionId),
      assignment.companyId,
    );
    if ("skip" in checked) return await skip(checked.skip);
    const contact = checked.contact;

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
          "The same email for this record was already sent in the last 10 minutes.",
        recipient: contact.email,
      };
    }

    const brand = await loadEmailBrandAttachments({
      companyName: input.record.companyName,
    });
    const template = buildEmail({
      action,
      register: input.register,
      record: input.record,
      category,
      actor,
      tmName: contact.displayName,
      includeCompanyLogo: brand.includeCompanyLogo,
    });

    // sendNotification applies Settings -> Notification types and logs the
    // outcome (sent / failed / skipped / not configured) itself.
    const result = await sendNotification({
      type: "training_record_change",
      category: REGISTER_CATEGORIES[input.register],
      to: contact.email,
      subject: template.subject,
      text: template.text,
      html: template.html,
      companyName: input.record.companyName ?? null,
      itemId: input.record.id,
      dedupeKey,
      actorEmail: input.actorEmail,
      detail: `${registerLabel} training ${action}`,
      fromName: "PAVE Training",
      attachments: brand.attachments.length ? brand.attachments : undefined,
    });

    return {
      attempted: true,
      skipped: result.status !== "sent",
      skipReason:
        result.status === "sent" ? undefined : (result.errorMessage ?? undefined),
      recipient: contact.email,
      result,
    };
  } catch (error) {
    // Never throw: the caller must not fail the save because of a mail error.
    const message =
      error instanceof Error ? error.message : "Unknown notification error.";
    console.error(
      `[notifyTrainingRecordChange] ${input.register} ${input.action} #${input.record.id}: ${message}`,
    );
    await writeNotificationLog({
      type: "training_record_change",
      status: "failed",
      companyName: input.record.companyName ?? null,
      subject: `${registerLabel} training ${input.action}: ${input.record.candidateName}`,
      itemId: input.record.id,
      errorMessage: `Notification failed before sending: ${message}`,
      detail: `${registerLabel} training ${input.action}`,
      actorEmail: input.actorEmail,
    });
    return {
      attempted: true,
      skipped: true,
      skipReason: `Notification failed: ${message}`,
      result: {
        status: "failed",
        recipientEmail: "",
        subject: "",
        errorMessage: message,
        logged: true,
      },
    };
  }
}

/**
 * Pass/Fail result stored on a register row before an edit, so the PATCH
 * route can tell a NEW result from a re-save. `undefined` when it cannot be
 * read (the notifier then treats any Pass/Fail as new, as before).
 */
export async function readRegisterOutcome(
  register: Register,
  id: string,
): Promise<string | null | undefined> {
  try {
    const item = await getListItemByKey(register, id);
    if (!item) return undefined;
    const f = getSharePointFields(register) as {
      trainingOutcome?: string;
      status?: string;
    };
    const read = (key?: string) => {
      const value = key ? item.fields[key] : null;
      return typeof value === "string" && value.trim() ? value.trim() : null;
    };
    return read(f.trainingOutcome) ?? read(f.status);
  } catch {
    return undefined;
  }
}

/**
 * Pure rules for Training Manager training-record emails: who may be emailed,
 * whether a save is a new result, and what the admin is told. Dependency-free
 * (types only) so scripts can test it — see scripts/test-training-notifications.mjs.
 * trainingRecordNotificationService.ts does the I/O.
 */

import type { PermissionProfile } from "@/types/models";
import type { NotificationSendResult } from "@/types/notifications";

export type ChangeAction = "added" | "updated" | "confirmed";

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
  /** NPORS card / registration number from the register row. */
  nporsNumber?: string | null;
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

/**
 * What the admin is told after a save. `sent` only when Graph accepted the
 * email - never inferred from "we tried".
 */
export interface TrainingManagerEmailSummary {
  status: "sent" | "skipped" | "failed" | "not_configured";
  message: string;
}

export function summarizeTrainingManagerEmail(
  outcome: NotifyTrainingRecordResult,
): TrainingManagerEmailSummary {
  const status = outcome.result?.status;
  if (status === "sent") {
    return {
      status: "sent",
      message: `Training Manager emailed (${outcome.recipient}).`,
    };
  }
  if (status === "failed") {
    return {
      status: "failed",
      message: `Training Manager email${outcome.recipient ? ` to ${outcome.recipient}` : ""} FAILED: ${outcome.result?.errorMessage ?? "unknown error"}. It is in the notification log.`,
    };
  }
  if (status === "not_configured") {
    return {
      status: "not_configured",
      message:
        "Training Manager email not sent — the notification mailbox is not configured.",
    };
  }
  return {
    status: "skipped",
    message: `Training Manager not emailed: ${outcome.skipReason ?? outcome.result?.errorMessage ?? "skipped"}`,
  };
}

export function isConfirmedOutcome(outcome: string | null | undefined): boolean {
  const value = outcome?.trim().toLowerCase();
  return value === "pass" || value === "fail" || value === "confirmed";
}

/**
 * "confirmed" (the result email) only when a Pass/Fail result is NEW: saved
 * with the record, or changed by this edit. Re-saving a record whose result
 * is already Pass is an ordinary "updated", not a second result email.
 */
export function resolveAction(
  requested: "added" | "updated",
  record: TrainingRecordNotificationRecord,
  previousOutcome?: string | null,
): ChangeAction {
  const result =
    isConfirmedOutcome(record.trainingOutcome) ||
    record.status?.trim().toLowerCase() === "completed";
  if (!result) return requested;
  if (requested === "added") return "confirmed";
  if (previousOutcome === undefined) return "confirmed";
  const before = previousOutcome?.trim().toLowerCase() ?? "";
  const after = (record.trainingOutcome ?? record.status ?? "")
    .trim()
    .toLowerCase();
  return before === after ? "updated" : "confirmed";
}

/** Contact for the assigned Training Manager. */
export interface TrainingManagerContact {
  email: string;
  displayName: string | null;
}

/**
 * The assigned TM's contact, or why they must not be emailed. Respects:
 *  - status: inactive portal access -> no email;
 *  - company: a customer-side TM must belong to the candidate's company (a
 *    stale assignment must never leak one company's results to another);
 *    a full Admin (all companies) is allowed;
 *  - preference: "Customer notifications enabled" switched off on their
 *    Permissions row -> no email.
 */
export function checkTrainingManagerContact(
  profile: PermissionProfile | null,
  candidateCompanyId: string | null,
): { contact: TrainingManagerContact } | { skip: string } {
  if (!profile) {
    return {
      skip: "Assigned Training Manager's Permissions row is missing or incomplete (email, status, role or company).",
    };
  }
  if (profile.status.trim().toLowerCase() !== "active") {
    return {
      skip: `Assigned Training Manager ${profile.userEmail} is ${profile.status}, not Active.`,
    };
  }
  if (!profile.customerNotificationsEnabled) {
    return {
      skip: `Assigned Training Manager ${profile.userEmail} has notifications switched off on their Permissions row.`,
    };
  }
  if (profile.customerRole !== null) {
    if (!candidateCompanyId) {
      return {
        skip: "Candidate's Workforce row has no company, so the Training Manager's company could not be checked.",
      };
    }
    if (String(profile.companyId) !== String(candidateCompanyId)) {
      return {
        skip: `Assigned Training Manager ${profile.userEmail} belongs to a different company (${profile.companyDisplayName ?? profile.companyId}) — not emailed. Fix the Training Manager on the Workforce row.`,
      };
    }
  }
  return {
    contact: {
      email: profile.userEmail,
      displayName: profile.candidateScopeName,
    },
  };
}

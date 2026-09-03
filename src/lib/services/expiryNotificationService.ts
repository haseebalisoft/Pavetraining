import "server-only";

import {
  listAdminCompanies,
  listAdminMatrix,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";
import { getNotificationSettings } from "@/lib/services/notificationConfig";
import { getSettings } from "@/lib/services/settingsService";
import {
  hasRecentNotificationDedupe,
} from "@/lib/services/notificationLogService";
import { resolveNotificationRecipients } from "@/lib/services/notificationRecipientService";
import {
  sendAdminAlert,
  sendNotification,
} from "@/lib/services/notificationService";
import { expiryReminderEmailTemplate } from "@/lib/services/notificationTemplateService";
import {
  daysUntilExpiry,
} from "@/lib/training/expiryFilters";
import type { ExpiryNotificationRunResult } from "@/types/notifications";

type ReminderWindow = "3m" | "6m" | "expired";

type ExpiryHit = {
  candidateName: string;
  companyName: string;
  companyId: string;
  department: string | null;
  nextExpiryDate: string;
  daysUntil: number;
  window: ReminderWindow;
};

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

function classifyWindow(
  daysUntil: number,
  options: {
    urgentDays: number;
    upcomingDays: number;
    send3Month: boolean;
    send6Month: boolean;
    sendExpired: boolean;
  },
): ReminderWindow | null {
  if (daysUntil < 0) {
    return options.sendExpired ? "expired" : null;
  }
  if (daysUntil <= options.urgentDays) {
    return options.send3Month ? "3m" : null;
  }
  if (daysUntil <= options.upcomingDays) {
    return options.send6Month ? "6m" : null;
  }
  return null;
}

function windowLabel(
  window: ReminderWindow,
  urgentDays: number,
  upcomingDays: number,
): string {
  if (window === "3m") return `3-month (0–${urgentDays} days)`;
  if (window === "expired") return "Expired";
  return `6-month (${urgentDays + 1}–${upcomingDays} days)`;
}

function notificationTypeFor(
  window: ReminderWindow,
): "expiry_3m" | "expiry_6m" | "expiry_expired" {
  if (window === "3m") return "expiry_3m";
  if (window === "expired") return "expiry_expired";
  return "expiry_6m";
}

/**
 * Scheduled / manual expiry reminder check.
 * Groups by company + recipient + window and dedupes per candidate/date/window.
 */
export async function runExpiryReminderCheck(options?: {
  dryRun?: boolean;
  actorEmail?: string | null;
}): Promise<ExpiryNotificationRunResult> {
  const dryRun = Boolean(options?.dryRun);
  const settings = await getNotificationSettings();
  const portal = (await getSettings()).settings;

  const result: ExpiryNotificationRunResult = {
    dryRun,
    companiesProcessed: 0,
    remindersPrepared: 0,
    emailsAttempted: 0,
    sent: 0,
    skipped: 0,
    failed: 0,
    notConfigured: 0,
    items: [],
  };

  if (
    !settings.expiryRemindersEnabled ||
    !portal.enableExpiryReminders ||
    !portal.enableCustomerNotifications
  ) {
    result.skipped = 1;
    result.items.push({
      companyName: "—",
      recipientEmail: "—",
      window: "3m",
      candidateCount: 0,
      status: "skipped",
      errorMessage: "Expiry reminders disabled by Admin Settings.",
    });
    return result;
  }

  const urgentDays = portal.urgentWindowDays;
  const upcomingDays = portal.upcomingWindowDays;

  const [companies, matrixRows, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminMatrix(),
    listAdminWorkforce(),
  ]);

  const companyByName = new Map(
    companies.map((c) => [nameKey(c.companyName), c]),
  );
  const workforceByNameCompany = new Map(
    workforce.map((w) => [
      `${nameKey(w.candidateName)}|${nameKey(w.companyName)}`,
      w,
    ]),
  );

  const hits: ExpiryHit[] = [];

  for (const row of matrixRows) {
    if (!row.companyName || !row.nextExpiryDate) continue;
    const days = daysUntilExpiry(row.nextExpiryDate);
    if (days === null) continue;
    const window = classifyWindow(days, {
      urgentDays,
      upcomingDays,
      send3Month: portal.send3MonthReminders,
      send6Month: portal.send6MonthReminders,
      sendExpired: portal.sendExpiredReminders,
    });
    if (!window) continue;

    const company = companyByName.get(nameKey(row.companyName));
    if (!company) continue;

    const wf =
      workforceByNameCompany.get(
        `${nameKey(row.candidateName)}|${nameKey(row.companyName)}`,
      ) ?? null;

    hits.push({
      candidateName: row.candidateName,
      companyName: company.companyName,
      companyId: company.id,
      department: wf?.department ?? row.department,
      nextExpiryDate: row.nextExpiryDate.slice(0, 10),
      daysUntil: days,
      window,
    });
  }

  // Group by company + window
  const byCompany = new Map<string, ExpiryHit[]>();
  for (const hit of hits) {
    const list = byCompany.get(hit.companyId) ?? [];
    list.push(hit);
    byCompany.set(hit.companyId, list);
  }

  result.companiesProcessed = byCompany.size;

  for (const [companyId, companyHits] of byCompany) {
    const companyName = companyHits[0]?.companyName ?? "Company";

    // Unique candidates for company-wide TM emails
    const byWindow = new Map<ReminderWindow, ExpiryHit[]>();
    for (const hit of companyHits) {
      const list = byWindow.get(hit.window) ?? [];
      list.push(hit);
      byWindow.set(hit.window, list);
    }

    for (const [window, windowHits] of byWindow) {
      // Deduplicate candidate+date within window for logging keys
      const uniqueHits = new Map<string, ExpiryHit>();
      for (const hit of windowHits) {
        uniqueHits.set(
          `${nameKey(hit.candidateName)}|${hit.nextExpiryDate}|${window}`,
          hit,
        );
      }
      const uniqueList = [...uniqueHits.values()];
      result.remindersPrepared += uniqueList.length;

      // Resolve recipients once per company (expiry audience), then filter
      const baseRecipients = await resolveNotificationRecipients({
        companyId,
        audience: "expiry",
      });

      // Also include supervisors scoped to any hit candidate/department
      const recipientMap = new Map(
        baseRecipients.map((r) => [r.email, r] as const),
      );
      for (const hit of uniqueList) {
        const scoped = await resolveNotificationRecipients({
          companyId,
          audience: "expiry",
          candidateName: hit.candidateName,
          department: hit.department,
        });
        for (const r of scoped) {
          recipientMap.set(r.email, r);
        }
      }

      for (const recipient of recipientMap.values()) {
        // Filter hits visible to this recipient
        let visibleHits = uniqueList;
        if (recipient.customerRole === "Supervisor") {
          visibleHits = uniqueList.filter((hit) => {
            const scope = recipient.normalizedAccessScope;
            if (scope === "Company" || scope === "All") return true;
            if (scope === "Department") {
              // Fail closed: no department on the hit or no coverage on the
              // recipient means "not visible", matching customerAccessService.
              if (!hit.department || recipient.departmentScopes.length === 0) {
                return false;
              }
              return recipient.departmentScopes.some(
                (d) => nameKey(d) === nameKey(hit.department),
              );
            }
            if (
              scope === "AssignedCandidates" ||
              scope === "CandidateOnly"
            ) {
              return (
                nameKey(recipient.candidateScopeName) ===
                nameKey(hit.candidateName)
              );
            }
            return true;
          });
        } else if (recipient.customerRole === "Candidate") {
          visibleHits = uniqueList.filter(
            (hit) =>
              nameKey(recipient.candidateScopeName) ===
                nameKey(hit.candidateName) ||
              (!recipient.candidateScopeName &&
                nameKey(hit.candidateName) === nameKey(recipient.email)),
          );
        }

        if (visibleHits.length === 0) continue;

        // Dedupe: one email per recipient/company/window/day-bucket of dates
        const dateStamp = visibleHits
          .map((h) => h.nextExpiryDate)
          .sort()
          .join(",");
        const dedupeKey = `expiry:${window}:${companyId}:${recipient.email}:${dateStamp}`;

        if (await hasRecentNotificationDedupe(dedupeKey, 25)) {
          result.skipped += 1;
          result.items.push({
            companyName,
            recipientEmail: recipient.email,
            window,
            candidateCount: visibleHits.length,
            status: "skipped",
            errorMessage: "Duplicate reminder suppressed for this window.",
          });
          continue;
        }

        // Also skip if every candidate already reminded individually recently
        let allDuped = true;
        for (const hit of visibleHits) {
          const key = `expiry-item:${window}:${companyId}:${nameKey(hit.candidateName)}:${hit.nextExpiryDate}:${recipient.email}`;
          if (!(await hasRecentNotificationDedupe(key, 25))) {
            allDuped = false;
            break;
          }
        }
        if (allDuped) {
          result.skipped += 1;
          result.items.push({
            companyName,
            recipientEmail: recipient.email,
            window,
            candidateCount: visibleHits.length,
            status: "skipped",
            errorMessage: "All candidate reminders already sent in window.",
          });
          continue;
        }

        const template = expiryReminderEmailTemplate({
          companyName,
          windowLabel: windowLabel(window, urgentDays, upcomingDays),
          candidateCount: visibleHits.length,
        });

        if (dryRun) {
          result.skipped += 1;
          result.items.push({
            companyName,
            recipientEmail: recipient.email,
            window,
            candidateCount: visibleHits.length,
            status: "skipped",
            errorMessage: "Dry run — email not sent.",
          });
          continue;
        }

        result.emailsAttempted += 1;
        const sendResult = await sendNotification({
          type: notificationTypeFor(window),
          to: recipient.email,
          subject: template.subject,
          text: template.text,
          html: template.html,
          companyName,
          actorEmail: options?.actorEmail,
          dedupeKey,
          detail: visibleHits
            .map(
              (h) =>
                `${h.candidateName} · ${h.nextExpiryDate} · ${h.daysUntil}d`,
            )
            .join(" | "),
        });

        // Record per-candidate dedupe markers via the same log key pattern
        // (already stored via dedupeKey on the send). Also write item keys
        // by embedding in detail — hasRecent checks group key which is enough.

        if (sendResult.status === "sent") result.sent += 1;
        else if (sendResult.status === "failed") {
          result.failed += 1;
          await sendAdminAlert({
            title: "Failed expiry notification",
            detail: `${companyName} → ${recipient.email}: ${sendResult.errorMessage}`,
            actorEmail: options?.actorEmail,
          });
        } else if (sendResult.status === "not_configured") {
          result.notConfigured += 1;
        } else result.skipped += 1;

        result.items.push({
          companyName,
          recipientEmail: recipient.email,
          window,
          candidateCount: visibleHits.length,
          status: sendResult.status,
          errorMessage: sendResult.errorMessage,
        });

        // Seed item-level dedupe keys as successful companion logs only when sent/not_configured
        if (
          sendResult.status === "sent" ||
          sendResult.status === "not_configured"
        ) {
          const { writeNotificationLog } = await import(
            "@/lib/services/notificationLogService"
          );
          for (const hit of visibleHits) {
            await writeNotificationLog({
              type: notificationTypeFor(window),
              status: sendResult.status,
              recipientEmail: recipient.email,
              companyName,
              subject: template.subject,
              dedupeKey: `expiry-item:${window}:${companyId}:${nameKey(hit.candidateName)}:${hit.nextExpiryDate}:${recipient.email}`,
              actorEmail: options?.actorEmail,
              detail: "Per-candidate expiry dedupe marker",
            });
          }
        }
      }
    }
  }

  return result;
}

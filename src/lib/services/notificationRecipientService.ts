import "server-only";

import { getListItemsByKey } from "@/lib/services/sharePointListService";
import {
  normalizePermissionRoleType,
  resolveCustomerRole,
  normalizeAccessScopeValue,
} from "@/lib/services/permissionService";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asLookupOrString,
  asString,
} from "@/lib/services/sharePointListService";
import type {
  CustomerRoleType,
  NormalizedAccessScope,
  PermissionProfile,
} from "@/types/models";

const permissionFields = getSharePointFields("permissions");

export type NotificationAudience =
  | "document"
  | "expiry"
  | "booking"
  | "admin_alert";

export interface NotificationRecipient {
  email: string;
  permissionId: string;
  companyId: string;
  companyName: string | null;
  /** Permissions.Name — used for email greetings. */
  displayName: string | null;
  customerRole: CustomerRoleType | null;
  roleLabel: string;
  normalizedAccessScope: NormalizedAccessScope;
  departmentScopes: string[];
  candidateScopeName: string | null;
}

function parseMultiChoice(value: unknown): string[] {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value
      .map((entry) => {
        if (typeof entry === "string") return entry.trim();
        if (entry && typeof entry === "object") {
          return (
            asString((entry as { LookupValue?: unknown }).LookupValue) ||
            asString((entry as { Title?: unknown }).Title) ||
            ""
          );
        }
        return "";
      })
      .filter(Boolean);
  }
  if (typeof value === "string") {
    return value
      .split(/;|#/)
      .map((part) => part.trim())
      .filter((part) => part && !/^\d+$/.test(part));
  }
  return [];
}

function preferenceOrDefault(
  fields: Record<string, unknown>,
  internal: string,
  defaultValue: boolean,
): boolean {
  if (!(internal in fields)) return defaultValue;
  return asBoolean(fields[internal]);
}

function mapActiveCustomerPermission(
  id: string,
  fields: Record<string, unknown>,
): PermissionProfile | null {
  const userEmail = asString(fields[permissionFields.userEmail])?.trim();
  const status = asString(fields[permissionFields.status])?.trim();
  const sharePointRoleType =
    asString(fields[permissionFields.roleType])?.trim() || "";
  const roleType = normalizePermissionRoleType(sharePointRoleType);
  const companyId =
    asString(fields[permissionFields.companyLookupId]) ??
    (typeof fields[permissionFields.company] === "object"
      ? asString(
          (fields[permissionFields.company] as { LookupId?: unknown }).LookupId,
        )
      : asString(fields[permissionFields.company]));
  const rawAccessScope = asString(fields[permissionFields.accessScope])?.trim();

  if (!userEmail || !status || !roleType || !companyId) return null;
  if (status.toLowerCase() !== "active") return null;

  const customerRole = resolveCustomerRole(sharePointRoleType, rawAccessScope ?? "");
  if (!customerRole) return null;

  // Same rule as permissionService.ts mapPermissionItem: a blank Access
  // Scope must not widen a Candidate role to company-wide notifications.
  const accessScope =
    rawAccessScope || (customerRole === "Candidate" ? "Candidate Only" : "Full Company");

  const departmentScopes = Array.from(
    new Set([
      ...parseMultiChoice(fields[permissionFields.departments]),
      ...parseMultiChoice(fields[permissionFields.departmentsAllowed]),
    ]),
  );

  return {
    id,
    userEmail: userEmail.toLowerCase(),
    status,
    roleType,
    sharePointRoleType,
    customerRole,
    roleLabel:
      customerRole === "TrainingManager"
        ? "Training Manager"
        : customerRole === "Supervisor"
          ? "Customer"
          : "Candidate",
    companyId,
    companyDisplayName:
      asLookupOrString(fields[permissionFields.company]) ?? undefined,
    accessScope,
    normalizedAccessScope: normalizeAccessScopeValue(
      accessScope,
      customerRole,
      false,
    ),
    departmentScopes,
    candidateScopeName: asString(fields[permissionFields.name])?.trim() || null,
    canView: asBoolean(fields[permissionFields.canView]),
    canDownload: asBoolean(fields[permissionFields.canDownload]),
    canEdit: asBoolean(fields[permissionFields.canEdit]),
    // Local to notification recipient shaping only — this is NOT the
    // authoritative auth flag. Notification code marks TMs as "admin
    // recipients" so they receive expiry / document emails for their
    // company, even though under the strict auth model TMs cannot access
    // /admin. Do not consume this field for portal access decisions.
    canAccessAdmin: customerRole === "TrainingManager",
    canAccessCustomer: true,
    receiveDocumentNotifications: preferenceOrDefault(
      fields,
      permissionFields.receiveDocumentNotifications,
      true,
    ),
    receiveExpiryNotifications: preferenceOrDefault(
      fields,
      permissionFields.receiveExpiryNotifications,
      true,
    ),
    customerNotificationsEnabled: preferenceOrDefault(
      fields,
      permissionFields.customerNotificationsEnabled,
      true,
    ),
  };
}

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

function supervisorRelevant(
  permission: PermissionProfile,
  context: {
    candidateName?: string | null;
    department?: string | null;
  },
): boolean {
  const scope = permission.normalizedAccessScope;
  if (scope === "Company" || scope === "All") return true;

  if (scope === "Department") {
    // Fail closed to match the in-app visibility rule (customerAccessService's
    // candidateRecordAllowed): no department coverage or no candidate
    // department means "not visible", not "visible to everyone".
    if (!context.department?.trim()) return false;
    if (permission.departmentScopes.length === 0) return false;
    const dept = nameKey(context.department);
    return permission.departmentScopes.some((d) => nameKey(d) === dept);
  }

  if (scope === "AssignedCandidates" || scope === "CandidateOnly") {
    if (!context.candidateName?.trim()) return false;
    if (!permission.candidateScopeName?.trim()) return false;
    return (
      nameKey(permission.candidateScopeName) === nameKey(context.candidateName)
    );
  }

  return true;
}

function candidateRelevant(
  permission: PermissionProfile,
  context: { candidateName?: string | null },
): boolean {
  if (!context.candidateName?.trim()) return false;
  if (permission.candidateScopeName?.trim()) {
    return (
      nameKey(permission.candidateScopeName) === nameKey(context.candidateName)
    );
  }
  // Fallback: role is Candidate with company match already applied.
  return permission.customerRole === "Candidate";
}

/**
 * Resolves Active customer-portal users who should receive a notification.
 */
export async function resolveNotificationRecipients(input: {
  companyId: string;
  audience: NotificationAudience;
  candidateName?: string | null;
  department?: string | null;
}): Promise<NotificationRecipient[]> {
  const items = await getListItemsByKey("permissions", { top: 5000 });
  const recipients: NotificationRecipient[] = [];
  const seen = new Set<string>();

  for (const item of items) {
    const permission = mapActiveCustomerPermission(item.id, item.fields);
    if (!permission) continue;
    if (permission.companyId !== input.companyId) continue;
    if (!permission.customerNotificationsEnabled) continue;
    if (!permission.canView) continue;

    if (input.audience === "document") {
      if (!permission.receiveDocumentNotifications) continue;
    }
    if (input.audience === "expiry") {
      if (!permission.receiveExpiryNotifications) continue;
    }
    // Booking confirmations: master customerNotificationsEnabled + Active TM
    // only (do not reuse document opt-out).

    const role = permission.customerRole;
    if (role === "TrainingManager") {
      // always include for company
    } else if (role === "Supervisor") {
      if (
        !supervisorRelevant(permission, {
          candidateName: input.candidateName,
          department: input.department,
        })
      ) {
        continue;
      }
    } else if (role === "Candidate") {
      if (input.audience === "expiry") {
        // Candidates generally get matrix via portal; only notify if candidate-specific
        if (
          !candidateRelevant(permission, {
            candidateName: input.candidateName,
          })
        ) {
          continue;
        }
      } else if (
        !candidateRelevant(permission, { candidateName: input.candidateName })
      ) {
        continue;
      }
    } else {
      continue;
    }

    if (seen.has(permission.userEmail)) continue;
    seen.add(permission.userEmail);

    recipients.push({
      email: permission.userEmail,
      permissionId: permission.id,
      companyId: permission.companyId,
      companyName: permission.companyDisplayName ?? null,
      displayName: permission.candidateScopeName,
      customerRole: permission.customerRole,
      roleLabel: permission.roleLabel,
      normalizedAccessScope: permission.normalizedAccessScope,
      departmentScopes: permission.departmentScopes,
      candidateScopeName: permission.candidateScopeName,
    });
  }

  return recipients;
}

/**
 * Active Training Managers for a company (booking confirmation audience).
 * Uses booking audience so document-upload opt-out does not suppress these.
 */
export async function resolveTrainingManagerRecipients(
  companyId: string,
): Promise<NotificationRecipient[]> {
  const all = await resolveNotificationRecipients({
    companyId,
    audience: "booking",
  });
  return all.filter(
    (recipient) => recipient.customerRole === "TrainingManager",
  );
}

/** Active Admin / PAVE staff emails for system alerts. */
export async function resolveAdminAlertRecipients(): Promise<string[]> {
  const items = await getListItemsByKey("permissions", { top: 5000 });
  const emails = new Set<string>();

  for (const item of items) {
    const userEmail = asString(item.fields[permissionFields.userEmail])
      ?.trim()
      .toLowerCase();
    const status = asString(item.fields[permissionFields.status])?.trim();
    const roleType = normalizePermissionRoleType(
      item.fields[permissionFields.roleType],
    );
    const sharePointRole =
      asString(item.fields[permissionFields.roleType])?.trim() || "";
    if (!userEmail || !status || status.toLowerCase() !== "active") continue;

    const customerRole = resolveCustomerRole(sharePointRole, "Full Company");
    // Admin bucket: RoleType Admin/Training Manager without customer-only use
    if (roleType === "Admin" && customerRole === null) {
      emails.add(userEmail);
      continue;
    }
    if (customerRole === "TrainingManager" && roleType === "Admin") {
      emails.add(userEmail);
    }
  }

  const override = process.env.NOTIFICATION_ADMIN_EMAILS?.trim();
  if (override) {
    for (const part of override.split(/[,;]+/)) {
      const email = part.trim().toLowerCase();
      if (email.includes("@")) emails.add(email);
    }
  }

  return [...emails];
}

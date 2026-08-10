import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import type { SharePointListKey } from "@/lib/schema/sharepointSchema";
import {
  ValidationError,
} from "@/lib/services/errorHandler";
import {
  buildFieldLookupIdEqualsFilter,
  deleteListItemByKey,
  extractLookupId,
  getListItemsByKey,
  updateListItemFieldsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";

const workforceFields = getSharePointFields("workforce");
const permissionFields = getSharePointFields("permissions");

export type InboundLookupClearTarget = {
  listKey: SharePointListKey;
  /** Graph `*LookupId` field name(s) that point at the item being deleted. */
  lookupIdFields: string[];
  /** Multi-lookup: remove id from array instead of nulling the whole field. */
  multi?: boolean;
};

/**
 * Pure cascade order used by company delete — exported for tests.
 * Related register/docs/matrix rows first, then workforce (clears Candidate
 * refs), then departments (clears Workforce/Permissions refs), then permissions.
 */
export const COMPANY_SAFE_DELETE_ORDER = [
  "customerDocuments",
  "trainingMatrix",
  "trainingMatrixCategoryRecords",
  "nporsRegister",
  "eusrRegister",
  "nrswaRegister",
  "inHouseCertificates",
  "nvqRegister",
  "events",
  "workforce",
  "departments",
  "permissions",
] as const satisfies ReadonlyArray<SharePointListKey>;

/** Inbound list refs that block Workforce delete (Restrict Delete). */
export const WORKFORCE_INBOUND_LOOKUPS: InboundLookupClearTarget[] = [
  { listKey: "nporsRegister", lookupIdFields: ["CandidateNameLookupId"] },
  { listKey: "eusrRegister", lookupIdFields: ["CandidateNameLookupId"] },
  { listKey: "nrswaRegister", lookupIdFields: ["CandidateNameLookupId"] },
  {
    listKey: "inHouseCertificates",
    lookupIdFields: ["CandidateNameLookupId"],
  },
  { listKey: "nvqRegister", lookupIdFields: ["CandidateNameLookupId"] },
  { listKey: "customerDocuments", lookupIdFields: ["CandidateLookupId"] },
  { listKey: "trainingMatrix", lookupIdFields: ["CandidateNameLookupId"] },
  {
    listKey: "trainingMatrixCategoryRecords",
    lookupIdFields: ["Candidate_x0020_NameLookupId"],
  },
  // "Training Matrix Update" also has a Workforce lookup. The matched row is
  // normally deleted first, but when the link is ambiguous the row is kept on
  // purpose — without this the leftover row would block the candidate delete.
  {
    listKey: "trainingMatrixExample",
    lookupIdFields: ["WorkforceLookupId"],
  },
];

/** Inbound list refs that block Department delete. */
export const DEPARTMENT_INBOUND_LOOKUPS: InboundLookupClearTarget[] = [
  {
    listKey: "workforce",
    lookupIdFields: [`${workforceFields.departmentText}LookupId`],
  },
  {
    listKey: "permissions",
    lookupIdFields: [`${permissionFields.departmentsAllowed}LookupId`],
    multi: true,
  },
];

/** Inbound list refs that block Permissions delete. */
export const PERMISSION_INBOUND_LOOKUPS: InboundLookupClearTarget[] = [
  {
    listKey: "workforce",
    lookupIdFields: [
      `${workforceFields.trainingManager}LookupId`,
      `${workforceFields.supervisor}LookupId`,
    ],
  },
];

function idsMatch(a: string | number | null | undefined, b: string): boolean {
  if (a == null || a === "") return false;
  return String(a).trim() === b || Number(a) === Number(b);
}

function extractMultiLookupIds(
  fields: SharePointFields,
  fieldInternalName: string,
): string[] {
  const sibling = fields[`${fieldInternalName}LookupId`];
  if (Array.isArray(sibling)) {
    return sibling.map((entry) => String(entry ?? "").trim()).filter(Boolean);
  }
  if (typeof sibling === "number" || typeof sibling === "string") {
    const text = String(sibling).trim();
    return text ? [text] : [];
  }
  const direct = fields[fieldInternalName];
  if (!Array.isArray(direct)) return [];
  return direct
    .map((entry) => {
      if (entry && typeof entry === "object" && "LookupId" in entry) {
        const id = (entry as { LookupId?: unknown }).LookupId;
        return id == null ? "" : String(id).trim();
      }
      return "";
    })
    .filter(Boolean);
}

/**
 * Null out (or strip from multi-lookup) every inbound ref so SharePoint
 * Restrict Delete allows removing the target item. Best-effort per row;
 * continues on individual patch failures and logs warnings.
 */
export async function clearInboundLookupRefs(
  targetItemId: string,
  targets: InboundLookupClearTarget[],
): Promise<number> {
  const trimmed = String(targetItemId ?? "").trim();
  const numericId = Number(trimmed);
  if (!trimmed || !Number.isFinite(numericId) || numericId <= 0) return 0;

  let cleared = 0;

  await Promise.all(
    targets.flatMap((target) =>
      target.lookupIdFields.map(async (lookupIdField) => {
        const fieldInternal = lookupIdField.replace(/LookupId$/, "");

        if (target.multi) {
          try {
            // Multi-lookup OData filters are unreliable — scan and strip.
            const items = await getListItemsByKey(target.listKey, { top: 5000 });
            await Promise.all(
              items.map(async (item) => {
                const ids = extractMultiLookupIds(item.fields, fieldInternal);
                if (!ids.some((id) => idsMatch(id, trimmed))) return;
                const next = ids
                  .filter((id) => !idsMatch(id, trimmed))
                  .map((id) => Number(id))
                  .filter((id) => Number.isFinite(id));
                try {
                  await updateListItemFieldsByKey(
                    target.listKey,
                    item.id,
                    {
                      [lookupIdField]: next.length > 0 ? next : null,
                    },
                    { skipReload: true },
                  );
                  cleared += 1;
                } catch (error) {
                  console.warn(
                    `[safeDelete] clear multi ${target.listKey}#${item.id} ${lookupIdField}:`,
                    error instanceof Error ? error.message : error,
                  );
                }
              }),
            );
          } catch (error) {
            console.warn(
              `[safeDelete] multi scan ${target.listKey}.${lookupIdField}:`,
              error instanceof Error ? error.message : error,
            );
          }
          return;
        }

        try {
          const items = await getListItemsByKey(target.listKey, {
            filter: buildFieldLookupIdEqualsFilter(lookupIdField, numericId),
            top: 5000,
          });
          await Promise.all(
            items.map(async (item) => {
              try {
                await updateListItemFieldsByKey(
                  target.listKey,
                  item.id,
                  {
                    [lookupIdField]: null,
                  },
                  { skipReload: true },
                );
                cleared += 1;
              } catch (error) {
                console.warn(
                  `[safeDelete] clear ${target.listKey}#${item.id} ${lookupIdField}:`,
                  error instanceof Error ? error.message : error,
                );
              }
            }),
          );
        } catch (error) {
          // Filter unsupported — fall back to full scan for this field.
          console.warn(
            `[safeDelete] filter ${target.listKey}.${lookupIdField} failed, scanning:`,
            error instanceof Error ? error.message : error,
          );
          try {
            const items = await getListItemsByKey(target.listKey, {
              top: 5000,
            });
            await Promise.all(
              items.map(async (item) => {
                const id = extractLookupId(item.fields, fieldInternal);
                if (!idsMatch(id, trimmed)) return;
                try {
                  await updateListItemFieldsByKey(
                    target.listKey,
                    item.id,
                    {
                      [lookupIdField]: null,
                    },
                    { skipReload: true },
                  );
                  cleared += 1;
                } catch (patchError) {
                  console.warn(
                    `[safeDelete] scan-clear ${target.listKey}#${item.id}:`,
                    patchError instanceof Error
                      ? patchError.message
                      : patchError,
                  );
                }
              }),
            );
          } catch (scanError) {
            console.warn(
              `[safeDelete] scan ${target.listKey}.${lookupIdField}:`,
              scanError instanceof Error ? scanError.message : scanError,
            );
          }
        }
      }),
    ),
  );

  return cleared;
}

export async function clearInboundLookupsToWorkforce(
  workforceId: string,
): Promise<number> {
  return clearInboundLookupRefs(workforceId, WORKFORCE_INBOUND_LOOKUPS);
}

export async function clearInboundLookupsToDepartment(
  departmentId: string,
): Promise<number> {
  return clearInboundLookupRefs(departmentId, DEPARTMENT_INBOUND_LOOKUPS);
}

export async function clearInboundLookupsToPermission(
  permissionId: string,
): Promise<number> {
  return clearInboundLookupRefs(permissionId, PERMISSION_INBOUND_LOOKUPS);
}

/**
 * Delete a list item after clearing known inbound Restrict Delete lookups.
 * Used for “admin owns all lists” — callers should prefer domain deletes when
 * they also clean non-lookup side effects (matrix seeds, Outlook, etc.).
 */
/**
 * After a Restrict-Delete failure, report WHICH lists still point at the item so
 * the admin gets a name instead of "ask a Site Owner". Only runs on the error
 * path, and only over the known inbound table (about ten cheap filtered reads).
 */
export async function describeBlockingReferences(
  targetItemId: string,
  targets: InboundLookupClearTarget[],
): Promise<string[]> {
  const trimmed = String(targetItemId ?? "").trim();
  const numericId = Number(trimmed);
  if (!trimmed || !Number.isFinite(numericId) || numericId <= 0) return [];

  const found: string[] = [];
  await Promise.all(
    targets.flatMap((target) =>
      target.lookupIdFields.map(async (lookupIdField) => {
        const fieldInternal = lookupIdField.replace(/LookupId$/, "");
        try {
          const items = target.multi
            ? (await getListItemsByKey(target.listKey, { top: 5000 })).filter(
                (item) =>
                  extractMultiLookupIds(item.fields, fieldInternal).some((id) =>
                    idsMatch(id, trimmed),
                  ),
              )
            : await getListItemsByKey(target.listKey, {
                filter: buildFieldLookupIdEqualsFilter(lookupIdField, numericId),
                top: 5000,
              });
          if (items.length) {
            const rows = items
              .slice(0, 5)
              .map((item) => `#${item.id}`)
              .join(", ");
            found.push(
              `${target.listKey} → ${fieldInternal} (${items.length} row${items.length === 1 ? "" : "s"}: ${rows}${items.length > 5 ? ", …" : ""})`,
            );
          }
        } catch {
          // Diagnostics only — a failed probe must not mask the delete error.
        }
      }),
    ),
  );
  return found.sort();
}

export async function safeDeleteListItem(
  listKey: SharePointListKey,
  itemId: string,
  options?: {
    inbound?: InboundLookupClearTarget[];
    label?: string;
  },
): Promise<void> {
  const trimmed = String(itemId ?? "").trim();
  if (!trimmed) {
    throw new ValidationError("Item id is required.");
  }

  if (options?.inbound?.length) {
    await clearInboundLookupRefs(trimmed, options.inbound);
  }

  try {
    await deleteListItemByKey(listKey, trimmed);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const related =
      /related to another item|cannot be deleted because|restrict/i.test(
        message,
      );
    const label = options?.label ?? listKey;
    if (related) {
      // Name the offending list(s) instead of sending the admin to a Site Owner.
      const blockers = options?.inbound?.length
        ? await describeBlockingReferences(trimmed, options.inbound)
        : [];
      throw new ValidationError(
        blockers.length
          ? `SharePoint blocks deleting this ${label} because these rows still reference it: ${blockers.join("; ")}. Remove or reassign them, then delete again.`
          : `SharePoint blocks deleting this ${label} because another list references it, and the reference is outside the lists this app manages. Ask a Site Owner which list has Restrict Delete pointing at ${listKey} item #${trimmed}.`,
      );
    }
    throw new ValidationError(
      message.includes("SharePoint") || message.includes("delete")
        ? message
        : `Could not delete this ${label}. ${message}`,
    );
  }
}

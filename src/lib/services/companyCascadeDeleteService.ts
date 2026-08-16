import "server-only";

import { getSharePointListId } from "@/lib/config/sharepoint";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import type { SharePointListKey } from "@/lib/schema/sharepointSchema";
import {
  asString,
  deleteListItemByKey,
  extractLookupId,
  getListItemById,
  getListItemByKey,
  getListItems,
} from "@/lib/services/sharePointListService";

const companyFields = getSharePointFields("company");

/**
 * Graph filter companions for Company List lookups.
 * Delete related rows before the company so SharePoint does not block.
 *
 * Order: unlink/delete child rows → workforce (clears Candidate refs) →
 * departments (clears Workforce/Permissions refs) → permissions → company.
 */
const COMPANY_CASCADE_TARGETS: ReadonlyArray<{
  listKey: SharePointListKey;
  lookupIdFields: string[];
  /** Base field name(s) for extractLookupId fallback when *LookupId sibling missing. */
  lookupDisplayFields?: string[];
  label: string;
  deleteMode?: "raw" | "workforce" | "department" | "permission";
}> = [
  {
    listKey: "customerDocuments",
    lookupIdFields: ["CompanyLookupId"],
    lookupDisplayFields: ["Company"],
    label: "Customer Documents",
  },
  {
    listKey: "trainingMatrix",
    lookupIdFields: ["MatrixCompanyLookupId", "Company_x0020_NameLookupId"],
    lookupDisplayFields: ["MatrixCompany", "Company_x0020_Name"],
    label: "Training Matrix",
  },
  {
    listKey: "trainingMatrixCategoryRecords",
    lookupIdFields: ["Company_x0020_NameLookupId"],
    lookupDisplayFields: ["Company_x0020_Name"],
    label: "Training Matrix Category Records",
  },
  {
    listKey: "nporsRegister",
    lookupIdFields: ["CompanyNameLookupId"],
    lookupDisplayFields: ["CompanyName"],
    label: "NPORS Register",
  },
  {
    listKey: "eusrRegister",
    lookupIdFields: ["CompanyNameLookupId"],
    lookupDisplayFields: ["CompanyName"],
    label: "EUSR Register",
  },
  {
    listKey: "nrswaRegister",
    lookupIdFields: ["CompanyNameLookupId"],
    lookupDisplayFields: ["CompanyName"],
    label: "Streetworks Register",
  },
  {
    listKey: "inHouseCertificates",
    lookupIdFields: ["CompanyNameLookupId"],
    lookupDisplayFields: ["CompanyName"],
    label: "In-House Certificates",
  },
  {
    listKey: "nvqRegister",
    lookupIdFields: ["NVQCompanyLookupId", "Company_x0020_NameLookupId"],
    lookupDisplayFields: ["NVQCompany", "Company_x0020_Name"],
    label: "NVQ Register",
  },
  {
    listKey: "events",
    lookupIdFields: ["EventCompanyLookupId"],
    lookupDisplayFields: ["EventCompany"],
    label: "Events",
  },
  {
    listKey: "workforce",
    lookupIdFields: ["CompanyNameLookupId"],
    lookupDisplayFields: ["CompanyName"],
    label: "Workforce",
    deleteMode: "workforce",
  },
  {
    // Training Matrix Update ("trainingMatrixExample") carries its own real
    // "Company" Lookup column (sibling CompanyLookupId) as well as the plain
    // CompanyItemId Number column added for the Workforce<->Matrix link work.
    // Rows linked to a Workforce record are already cleared by the
    // "workforce" cascade step above (deleteAdminWorkforce clears/deletes its
    // matched matrix row); this step catches unlinked/Needs Review rows that
    // still reference this company directly and would otherwise leave a
    // SharePoint Restrict Delete block on the company.
    listKey: "trainingMatrixExample",
    lookupIdFields: ["CompanyLookupId", "CompanyItemId"],
    lookupDisplayFields: ["Company"],
    label: "Training Matrix Update",
  },
  {
    listKey: "departments",
    lookupIdFields: ["CompanyLookupId"],
    lookupDisplayFields: ["Company"],
    label: "Departments",
    deleteMode: "department",
  },
  {
    listKey: "permissions",
    lookupIdFields: ["CompanyLookupId"],
    lookupDisplayFields: ["Company"],
    label: "Permissions",
    deleteMode: "permission",
  },
];

export interface CompanyCascadeResult {
  companyId: string;
  companyName: string;
  relatedDeleted: number;
  companyDeleted: boolean;
  details: string[];
  errors: string[];
}

function idsMatchCompany(
  fields: Record<string, unknown>,
  companyId: string,
  lookupIdFields: string[],
  lookupDisplayFields?: string[],
): boolean {
  for (const field of lookupIdFields) {
    const raw = fields[field];
    if (raw != null && String(raw).trim() === companyId) return true;
  }
  for (const display of lookupDisplayFields ?? []) {
    const id = extractLookupId(fields, display);
    if (id && id === companyId) return true;
  }
  return false;
}

async function deleteCascadeItem(
  listKey: SharePointListKey,
  itemId: string,
  deleteMode: "raw" | "workforce" | "department" | "permission" | undefined,
): Promise<void> {
  if (deleteMode === "workforce") {
    const { deleteAdminWorkforce } = await import(
      "@/lib/services/adminCrudService"
    );
    await deleteAdminWorkforce(itemId);
    return;
  }
  if (deleteMode === "department") {
    const { deleteAdminDepartment } = await import(
      "@/lib/services/departmentService"
    );
    await deleteAdminDepartment(itemId);
    return;
  }
  if (deleteMode === "permission") {
    const { deleteAdminPermission } = await import(
      "@/lib/services/adminCrudService"
    );
    try {
      await deleteAdminPermission(itemId);
    } catch (error) {
      // Protected admins must not be wiped by company cascade — unlink company
      // instead so the company row can still be deleted.
      const message = error instanceof Error ? error.message : String(error);
      if (/protected|cannot be deleted|cannot delete/i.test(message)) {
        const { updateListItemFieldsByKey } = await import(
          "@/lib/services/sharePointListService"
        );
        await updateListItemFieldsByKey(
          "permissions",
          itemId,
          { CompanyLookupId: null },
          { skipReload: true },
        );
        return;
      }
      throw error;
    }
    return;
  }
  await deleteListItemByKey(listKey, itemId);
}

/**
 * Bounded concurrency helper — runs `worker` for each item, up to `concurrency`
 * in-flight at a time. Prevents SharePoint from throttling us when a company
 * has hundreds of related rows (e.g. large Workforce lists).
 */
async function runBounded<Item>(
  items: Item[],
  concurrency: number,
  worker: (item: Item) => Promise<void>,
): Promise<void> {
  if (items.length === 0) return;
  const cap = Math.max(1, Math.min(concurrency, items.length));
  let cursor = 0;
  const runners: Promise<void>[] = [];
  for (let i = 0; i < cap; i += 1) {
    runners.push(
      (async () => {
        while (true) {
          const index = cursor;
          cursor += 1;
          if (index >= items.length) return;
          await worker(items[index]!);
        }
      })(),
    );
  }
  await Promise.all(runners);
}

/**
 * How many item deletions to run in parallel inside one cascade step.
 * SharePoint tolerates this comfortably per-list; keeps the tail latency
 * of the Workforce / Departments / Permissions cascade short.
 */
const CASCADE_ITEM_CONCURRENCY = 6;

async function deleteByLookupId(
  listKey: SharePointListKey,
  lookupIdFields: string[],
  companyId: number,
  label: string,
  result: CompanyCascadeResult,
  deleteMode?: "raw" | "workforce" | "department" | "permission",
  lookupDisplayFields?: string[],
): Promise<void> {
  const companyIdText = String(companyId);
  const seen = new Set<string>();
  let deletedCount = 0;

  const deleteItem = async (itemId: string) => {
    try {
      await deleteCascadeItem(listKey, itemId, deleteMode);
      deletedCount += 1;
      result.relatedDeleted += 1;
    } catch (error) {
      result.errors.push(
        `${label} #${itemId}: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
  };

  for (const field of lookupIdFields) {
    try {
      // Uncached — company delete must see latest children.
      const items = await getListItems(getSharePointListId(listKey), {
        filter: `fields/${field} eq ${companyId}`,
        top: 5000,
      });

      const fresh: string[] = [];
      for (const item of items) {
        if (seen.has(item.id)) continue;
        seen.add(item.id);
        fresh.push(item.id);
      }
      await runBounded(fresh, CASCADE_ITEM_CONCURRENCY, deleteItem);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (!message.toLowerCase().includes("not found")) {
        result.errors.push(`${label} (${field}): ${message}`);
      }
    }
  }

  // Fallback: OData filter can miss LookupId siblings — scan once.
  if (seen.size === 0) {
    try {
      const all = await getListItems(getSharePointListId(listKey), {
        top: 5000,
      });
      const fresh: string[] = [];
      for (const item of all) {
        if (seen.has(item.id)) continue;
        if (
          !idsMatchCompany(
            item.fields,
            companyIdText,
            lookupIdFields,
            lookupDisplayFields,
          )
        ) {
          continue;
        }
        seen.add(item.id);
        fresh.push(item.id);
      }
      await runBounded(fresh, CASCADE_ITEM_CONCURRENCY, deleteItem);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (!message.toLowerCase().includes("not found")) {
        result.errors.push(`${label} (scan): ${message}`);
      }
    }
  }

  if (deletedCount > 0) {
    result.details.push(`${label}: deleted ${deletedCount} item(s)`);
  } else if (seen.size > 0) {
    result.details.push(
      `${label}: found ${seen.size} item(s) but none could be deleted`,
    );
  }
}

export async function deleteCompanyWithRelatedData(
  companyId: string,
): Promise<CompanyCascadeResult> {
  const numericId = Number(companyId);
  const result: CompanyCascadeResult = {
    companyId,
    companyName: "",
    relatedDeleted: 0,
    companyDeleted: false,
    details: [],
    errors: [],
  };

  if (!Number.isFinite(numericId) || numericId <= 0) {
    result.errors.push(`Invalid company id: ${companyId}`);
    return result;
  }

  const company = await getListItemByKey("company", companyId);
  result.companyName =
    asString(company?.fields[companyFields.companyName]) ||
    asString(company?.fields.Title) ||
    `#${companyId}`;

  // Speed: the leaf cascade targets are independent (different SharePoint
  // lists) and each `deleteByLookupId` awaits its own item deletions
  // sequentially. Running the leaf scans in parallel drops a real-world
  // delete from ~2 min to ~15–30 s. Rows are pushed to `result.details` /
  // `result.errors` inside each call — safe under Node's single-thread
  // event loop (no interleaved writes to the same array cell).
  //
  // We keep the CROSS-LINKED chain (workforce → trainingMatrixExample →
  // departments → permissions) SEQUENTIAL in the exact previous order so
  // ordering guarantees inside their per-entity deleters (deleteAdminWorkforce
  // clearing linked matrix rows, deleteAdminDepartment clearing Workforce/
  // Permissions refs, deleteAdminPermission's protected-admin unlink) are
  // preserved bit-for-bit.
  const sequentialListKeys = new Set([
    "workforce",
    "trainingMatrixExample",
    "departments",
    "permissions",
  ]);
  const parallelTargets = COMPANY_CASCADE_TARGETS.filter(
    (target) => !sequentialListKeys.has(target.listKey),
  );
  const sequentialTargets = COMPANY_CASCADE_TARGETS.filter((target) =>
    sequentialListKeys.has(target.listKey),
  );

  await Promise.all(
    parallelTargets.map((target) =>
      deleteByLookupId(
        target.listKey,
        target.lookupIdFields,
        numericId,
        target.label,
        result,
        target.deleteMode,
        target.lookupDisplayFields,
      ),
    ),
  );

  for (const target of sequentialTargets) {
    await deleteByLookupId(
      target.listKey,
      target.lookupIdFields,
      numericId,
      target.label,
      result,
      target.deleteMode,
      target.lookupDisplayFields,
    );
  }

  // Training Manager Logs are intentionally NOT cascade-deleted here — the
  // audit trail (including this delete's own entry) must survive the entity
  // it describes, per the retention policy in auditLogService.ts.

  // Always attempt company delete after cascade. Soft related errors should not
  // skip the final delete — SharePoint will reject if Restrict Delete remains.
  try {
    await deleteListItemByKey("company", companyId);
    // Verify the row is truly gone before reporting success. Some tenants/paths
    // return success while the item survives (SharePoint "Restrict Delete" on a
    // related list, or a recycle no-op), which the UI would show as a false
    // "deleted" that reappears on refresh. A direct (uncached) re-read is
    // strongly consistent right after a Graph delete; a read error is treated as
    // "gone" so a transient blip can't turn a real delete into a false failure.
    const survivor = await getListItemById(
      getSharePointListId("company"),
      companyId,
    ).catch(() => null);
    if (survivor) {
      result.companyDeleted = false;
      result.errors.push(
        "Company still exists after delete — SharePoint is blocking it because a related row still references it (Workforce, Departments, Permissions, Registers, Documents, or Events). Resolve the related warnings above and retry.",
      );
    } else {
      result.companyDeleted = true;
      result.details.push("Company deleted");
      if (result.errors.length > 0) {
        result.details.push(
          `Company deleted with ${result.errors.length} related warning(s).`,
        );
      }
    }
  } catch (error) {
    result.errors.push(
      `Company delete failed: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
    result.errors.push(
      "Related rows may still block SharePoint Restrict Delete (Workforce, Departments, Permissions, Registers, Documents, Events). Fix those errors and retry.",
    );
  }

  return result;
}

/**
 * Bulk delete concurrency: how many companies to cascade-delete at the same
 * time. Each single-company delete is already up to 6-concurrent at the item
 * level, so the effective peak in-flight Graph calls is roughly
 * `BULK_COMPANY_CONCURRENCY * CASCADE_ITEM_CONCURRENCY` (~12). SharePoint
 * handles that comfortably without throttling.
 */
const BULK_COMPANY_CONCURRENCY = 2;

export async function deleteCompaniesWithRelatedData(companyIds: string[]) {
  const orderedResults: CompanyCascadeResult[] = new Array(companyIds.length);
  let companiesDeleted = 0;
  let relatedDeleted = 0;

  await runBounded(
    companyIds.map((id, index) => ({ id, index })),
    BULK_COMPANY_CONCURRENCY,
    async ({ id, index }) => {
      const one = await deleteCompanyWithRelatedData(id);
      orderedResults[index] = one;
      if (one.companyDeleted) companiesDeleted += 1;
      relatedDeleted += one.relatedDeleted;
    },
  );

  return { companiesDeleted, relatedDeleted, results: orderedResults };
}

export const COMPANY_CASCADE_LABELS = COMPANY_CASCADE_TARGETS.map(
  (t) => t.label,
);

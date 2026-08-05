import "server-only";

import { getSharePointListId } from "@/lib/config/sharepoint";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import type { SharePointListKey } from "@/lib/schema/sharepointSchema";
import {
  asString,
  deleteListItemByKey,
  extractLookupId,
  getListItemByKey,
  getListItems,
  getListItemsByKey,
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

  for (const field of lookupIdFields) {
    try {
      // Uncached — company delete must see latest children.
      const items = await getListItems(getSharePointListId(listKey), {
        filter: `fields/${field} eq ${companyId}`,
        top: 5000,
      });

      for (const item of items) {
        if (seen.has(item.id)) continue;
        seen.add(item.id);
        try {
          await deleteCascadeItem(listKey, item.id, deleteMode);
          deletedCount += 1;
          result.relatedDeleted += 1;
        } catch (error) {
          result.errors.push(
            `${label} #${item.id}: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        }
      }
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
        try {
          await deleteCascadeItem(listKey, item.id, deleteMode);
          deletedCount += 1;
          result.relatedDeleted += 1;
        } catch (error) {
          result.errors.push(
            `${label} #${item.id}: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        }
      }
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

  for (const target of COMPANY_CASCADE_TARGETS) {
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

  if (result.companyName && !result.companyName.startsWith("#")) {
    try {
      const escaped = result.companyName.replace(/'/g, "''");
      const logs = await getListItemsByKey("trainingManagerLogs", {
        filter: `fields/Company eq '${escaped}'`,
        top: 5000,
      });
      let logDeleted = 0;
      for (const item of logs) {
        try {
          await deleteListItemByKey("trainingManagerLogs", item.id);
          result.relatedDeleted += 1;
          logDeleted += 1;
        } catch {
          // Logs must not block company delete.
        }
      }
      if (logDeleted > 0) {
        result.details.push(
          `Training Manager Logs: deleted ${logDeleted} item(s)`,
        );
      }
    } catch {
      // optional list / field
    }
  }

  // Always attempt company delete after cascade. Soft related errors should not
  // skip the final delete — SharePoint will reject if Restrict Delete remains.
  try {
    await deleteListItemByKey("company", companyId);
    result.companyDeleted = true;
    result.details.push("Company deleted");
    if (result.errors.length > 0) {
      result.details.push(
        `Company deleted with ${result.errors.length} related warning(s).`,
      );
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

export async function deleteCompaniesWithRelatedData(companyIds: string[]) {
  const results: CompanyCascadeResult[] = [];
  let companiesDeleted = 0;
  let relatedDeleted = 0;

  for (const id of companyIds) {
    const one = await deleteCompanyWithRelatedData(id);
    results.push(one);
    if (one.companyDeleted) companiesDeleted += 1;
    relatedDeleted += one.relatedDeleted;
  }

  return { companiesDeleted, relatedDeleted, results };
}

export const COMPANY_CASCADE_LABELS = COMPANY_CASCADE_TARGETS.map(
  (t) => t.label,
);

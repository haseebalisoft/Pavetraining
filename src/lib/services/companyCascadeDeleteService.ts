import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import type { SharePointListKey } from "@/lib/schema/sharepointSchema";
import {
  asString,
  deleteListItemByKey,
  getListItemByKey,
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
  label: string;
  deleteMode?: "raw" | "workforce" | "department" | "permission";
}> = [
  {
    listKey: "customerDocuments",
    lookupIdFields: ["CompanyLookupId"],
    label: "Customer Documents",
  },
  {
    listKey: "trainingMatrix",
    lookupIdFields: ["MatrixCompanyLookupId", "Company_x0020_NameLookupId"],
    label: "Training Matrix",
  },
  {
    listKey: "trainingMatrixCategoryRecords",
    lookupIdFields: ["Company_x0020_NameLookupId"],
    label: "Training Matrix Category Records",
  },
  {
    listKey: "nporsRegister",
    lookupIdFields: ["CompanyNameLookupId"],
    label: "NPORS Register",
  },
  {
    listKey: "eusrRegister",
    lookupIdFields: ["CompanyNameLookupId"],
    label: "EUSR Register",
  },
  {
    listKey: "nrswaRegister",
    lookupIdFields: ["CompanyNameLookupId"],
    label: "Streetworks Register",
  },
  {
    listKey: "inHouseCertificates",
    lookupIdFields: ["CompanyNameLookupId"],
    label: "In-House Certificates",
  },
  {
    listKey: "nvqRegister",
    lookupIdFields: ["NVQCompanyLookupId", "Company_x0020_NameLookupId"],
    label: "NVQ Register",
  },
  {
    listKey: "events",
    lookupIdFields: ["EventCompanyLookupId"],
    label: "Events",
  },
  {
    listKey: "workforce",
    lookupIdFields: ["CompanyNameLookupId"],
    label: "Workforce",
    deleteMode: "workforce",
  },
  {
    listKey: "departments",
    lookupIdFields: ["CompanyLookupId"],
    label: "Departments",
    deleteMode: "department",
  },
  {
    listKey: "permissions",
    lookupIdFields: ["CompanyLookupId"],
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
        await updateListItemFieldsByKey("permissions", itemId, {
          CompanyLookupId: null,
        });
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
): Promise<void> {
  const seen = new Set<string>();
  let deletedCount = 0;

  for (const field of lookupIdFields) {
    try {
      const items = await getListItemsByKey(listKey, {
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
    );
  }

  if (result.companyName && !result.companyName.startsWith("#")) {
    try {
      const escaped = result.companyName.replace(/'/g, "''");
      const logs = await getListItemsByKey("trainingManagerLogs", {
        filter: `fields/Company eq '${escaped}'`,
        top: 5000,
      });
      for (const item of logs) {
        try {
          await deleteListItemByKey("trainingManagerLogs", item.id);
          result.relatedDeleted += 1;
        } catch (error) {
          result.errors.push(
            `Logs #${item.id}: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        }
      }
      if (logs.length > 0) {
        result.details.push(
          `Training Manager Logs: deleted ${logs.length} item(s)`,
        );
      }
    } catch {
      // optional list / field
    }
  }

  if (result.errors.length > 0) {
    result.errors.push(
      "Company was not deleted because related rows still block SharePoint Restrict Delete. Retry — cascade clears Workforce / Department / Permission lookups before delete.",
    );
    return result;
  }

  try {
    await deleteListItemByKey("company", companyId);
    result.companyDeleted = true;
    result.details.push("Company deleted");
  } catch (error) {
    result.errors.push(
      `Company delete failed: ${
        error instanceof Error ? error.message : String(error)
      }`,
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

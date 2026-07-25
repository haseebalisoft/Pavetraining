import "server-only";

import { cache } from "react";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asNullableString,
  asString,
  getListItemByKey,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import type { Company, CustomerCompanyProfile } from "@/types/models";

const companyFields = getSharePointFields("company");

export function mapCompanyFields(
  id: string,
  fields: SharePointFields | Record<string, unknown>,
): Company | null {
  const title = asString(fields[companyFields.title]);
  const companyName =
    asString(fields[companyFields.companyName]) ?? title;

  if (!companyName) {
    return null;
  }

  return {
    id,
    title: title ?? companyName,
    companyName,
    companyNumber: asNullableString(fields[companyFields.companyNumber]),
    companySize: asNullableString(fields[companyFields.companySize]),
    registeredAddress: asNullableString(
      fields[companyFields.registeredAddress],
    ),
    companyRegNumber: asNullableString(
      fields[companyFields.companyRegNumber],
    ),
    vatNo: asNullableString(fields[companyFields.vatNo]),
    telNo: asNullableString(fields[companyFields.telNo]),
    email: asNullableString(fields[companyFields.email]),
    mainContact: asNullableString(fields[companyFields.mainContact]),
    accountsContactName: asNullableString(
      fields[companyFields.accountsContactName],
    ),
    accountsAddress: asNullableString(fields[companyFields.accountsAddress]),
    accountsContactNumber: asNullableString(
      fields[companyFields.accountsContactNumber],
    ),
    accountsEmail: asNullableString(fields[companyFields.accountsEmail]),
    notesPricesAgreed: asNullableString(
      fields[companyFields.notesPricesAgreed],
    ),
    companyLogo: asNullableString(fields[companyFields.companyLogo]),
    status: asNullableString(fields[companyFields.status]) ?? "Unknown",
  };
}

export function toCustomerCompanyProfile(
  company: Company,
): CustomerCompanyProfile {
  return {
    id: company.id,
    companyName: company.companyName,
    companyNumber: company.companyNumber,
    companySize: company.companySize,
    registeredAddress: company.registeredAddress,
    companyRegNumber: company.companyRegNumber,
    vatNo: company.vatNo,
    telNo: company.telNo,
    email: company.email,
    mainContact: company.mainContact,
    companyLogo: company.companyLogo,
    status: company.status,
  };
}

/**
 * Resolves company records from the SharePoint Company List.
 * Deduped per request via React.cache.
 */
export const getCompanyById = cache(
  async (companyId: string): Promise<Company | null> => {
    const trimmedId = companyId.trim();
    if (!trimmedId) {
      return null;
    }

    const item = await getListItemByKey("company", trimmedId);

    if (!item) {
      return null;
    }

    return mapCompanyFields(item.id, item.fields);
  },
);

/**
 * Lists all companies for admin portals. Admins may view every company.
 */
export async function getAllCompanies(): Promise<Company[]> {
  const items = await getListItemsByKey("company", { top: 5000 });

  return items
    .map((item) => mapCompanyFields(item.id, item.fields))
    .filter((company): company is Company => company !== null)
    .sort((a, b) => a.companyName.localeCompare(b.companyName));
}

export function isActiveCompanyStatus(
  status: string | null | undefined,
): boolean {
  const normalized = (status ?? "").trim().toLowerCase();
  return (
    normalized === "active" ||
    normalized === "live" ||
    normalized === "enabled" ||
    normalized === ""
  );
}

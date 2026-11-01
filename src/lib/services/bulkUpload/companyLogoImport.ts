import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  planImportImage,
  storeImportImageFromUrl,
} from "@/lib/services/bulkUpload/importImage";

const companyFields = getSharePointFields("company");

export type CompanyLogoPlan =
  /** Blank / placeholder: the logo is optional, so this is a note, not a warning. */
  | { kind: "missing"; note: string }
  | { kind: "invalid"; warning: string }
  | { kind: "url"; url: string; fileName: string }
  /** Image file name — the admin attaches the file on the upload page. */
  | { kind: "file"; fileName: string };

/**
 * Company Logo cells accept an https/http image URL, or the file name of a
 * logo image (e.g. "alder-finch.png") that the admin attaches under "Company
 * logo files" on the bulk upload page — the browser uploads those after the
 * import. Placeholders and blanks are treated as missing (note only — row
 * still imports). Free text is never written into the SharePoint Thumbnail
 * column.
 */
export function planCompanyLogoImport(
  value: string | null | undefined,
): CompanyLogoPlan {
  const plan = planImportImage(value);
  if (plan.kind === "missing") {
    return {
      kind: "missing",
      note: plan.placeholder
        ? `No logo ("${plan.placeholder}" is a placeholder) — the company's logo is left unchanged.`
        : "No logo in the spreadsheet — the company's logo is left unchanged.",
    };
  }
  if (plan.kind === "invalid") {
    return {
      kind: "invalid",
      warning: `Company Logo ${plan.reason} — company will import without a logo.`,
    };
  }
  return plan;
}

/**
 * After the Company List row exists, download the logo URL and store it with
 * the same PortalMedia + CompanyLogo Thumbnail path as Admin logo upload.
 * Returns a warning string on skip/failure (never throws for logo problems).
 */
export async function applyCompanyLogoFromImportCell(
  companyId: string,
  cell: string | null | undefined,
): Promise<string | null> {
  const plan = planCompanyLogoImport(cell);
  if (plan.kind === "missing") return null;
  if (plan.kind === "invalid") return plan.warning;
  // Attached file — the bulk upload page uploads it once the row exists.
  if (plan.kind === "file") return null;

  const failure = await storeImportImageFromUrl({
    listKey: "company",
    itemId: companyId,
    fieldInternalName: companyFields.companyLogo,
    url: plan.url,
    fileName: plan.fileName,
  });
  return failure
    ? `Company Logo upload failed (${failure}) — company imported without a logo.`
    : null;
}

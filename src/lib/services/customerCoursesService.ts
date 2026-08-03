import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asNullableString,
  getListItemsByKey,
} from "@/lib/services/sharePointListService";
const fields = getSharePointFields("nporsCategories");

export type CustomerCourseCatalogueItem = {
  code: string;
  title: string;
  group?: string | null;
};

function extractCode(title: string | null | undefined): string | null {
  if (!title?.trim()) return null;
  const match = title.trim().match(/^(N\d+[A-Z]?)\b/i);
  return match?.[1]?.toUpperCase() ?? null;
}

/** Customer-visible PAVE course catalogue from SharePoint NPORS Categories. */
export async function getCustomerCourseCatalogue(): Promise<
  CustomerCourseCatalogueItem[]
> {
  try {
    const items = await getListItemsByKey("nporsCategories", { top: 5000 });
    const courses: CustomerCourseCatalogueItem[] = [];
    const seen = new Set<string>();

    for (const item of items) {
      if (
        item.fields[fields.customerVisible] != null &&
        !asBoolean(item.fields[fields.customerVisible])
      ) {
        continue;
      }
      if (
        item.fields[fields.active] != null &&
        !asBoolean(item.fields[fields.active])
      ) {
        continue;
      }

      const title = asNullableString(item.fields[fields.title]);
      const code =
        extractCode(title) ??
        extractCode(asNullableString(item.fields[fields.categoryCode]));
      if (!code || seen.has(code)) continue;
      seen.add(code);

      const machine = asNullableString(item.fields[fields.machineType]);
      const displayTitle =
        title?.replace(new RegExp(`^${code}\\s*[-–—]\\s*`, "i"), "").trim() ||
        machine ||
        code;

      courses.push({
        code,
        title: displayTitle,
        group: asNullableString(item.fields[fields.categoryGroup]),
      });
    }

    courses.sort((a, b) =>
      a.code.localeCompare(b.code, undefined, { numeric: true }),
    );
    return courses;
  } catch (error) {
    console.warn("[customerCourses] Failed to load NPORS Categories", error);
    return [];
  }
}

import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asNullableString,
  getListItemsByKey,
} from "@/lib/services/sharePointListService";
import { getNporsCategoryOptions } from "@/lib/training/nporsCategoryOptions";

export type NporsCategoryOption = {
  value: string;
  label: string;
};

const fields = getSharePointFields("nporsCategories");

function extractCode(title: string | null | undefined): string | null {
  if (!title?.trim()) return null;
  const match = title.trim().match(/^(N\d+[A-Z]?)\b/i);
  return match?.[1]?.toUpperCase() ?? null;
}

/**
 * Live options from SharePoint list "NPORS Categories".
 * Label = Title (e.g. "N001 - Industrial Counterbalanced Lift Truck").
 * Value = code only (e.g. "N001") for NPORS Register MultiChoice + matrix sync.
 * Falls back to matrix template short list if the SharePoint list is empty.
 */
export async function listAdminNporsCategoryOptions(): Promise<
  NporsCategoryOption[]
> {
  try {
    const items = await getListItemsByKey("nporsCategories", { top: 5000 });
    const options: NporsCategoryOption[] = [];
    const seen = new Set<string>();

    for (const item of items) {
      const active = asBoolean(item.fields[fields.active]);
      // Treat missing Active as true (legacy rows).
      if (item.fields[fields.active] != null && !active) continue;

      const title =
        asNullableString(item.fields[fields.title]) ??
        asNullableString(item.fields[fields.categoryCode]);
      const code =
        extractCode(title) ??
        extractCode(asNullableString(item.fields[fields.categoryCode]));
      if (!code || seen.has(code)) continue;
      seen.add(code);

      const machine =
        asNullableString(item.fields[fields.machineType])?.trim() || null;
      const label =
        title?.trim() ||
        (machine ? `${code} - ${machine}` : code);

      // Skip incomplete code-only labels when a fuller title is expected.
      if (label.toUpperCase() === code && !machine) {
        continue;
      }

      options.push({ value: code, label });
    }

    options.sort((a, b) =>
      a.value.localeCompare(b.value, undefined, { numeric: true }),
    );

    if (options.length > 0) return options;
  } catch (error) {
    console.warn(
      "[nporsCategories] Failed to load SharePoint NPORS Categories list — using template fallback.",
      error,
    );
  }

  return getNporsCategoryOptions();
}

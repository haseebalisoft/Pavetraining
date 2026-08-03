import { CLIENT_MATRIX_CATEGORY_COLUMNS } from "@/lib/services/bulkUpload/clientTemplateHeaders";

const EXTRA_IN_HOUSE_COURSES = [
  "Face Fit",
  "Asbestos Awareness",
] as const;

/**
 * Admin In-House Course dropdown options:
 * all NPORS category display names + Face Fitting + Asbestos Awareness.
 */
export function getInHouseCourseOptions(): Array<{
  value: string;
  label: string;
}> {
  const seen = new Set<string>();
  const options: Array<{ value: string; label: string }> = [];

  for (const column of CLIENT_MATRIX_CATEGORY_COLUMNS) {
    // Prefer full SharePoint-style header (e.g. "N001 - Ind FLT").
    const label = column.header?.trim() || column.name?.trim();
    const key = label.toLowerCase();
    if (!label || seen.has(key)) continue;
    seen.add(key);
    options.push({ value: label, label });
  }

  for (const label of EXTRA_IN_HOUSE_COURSES) {
    const key = label.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    options.push({ value: label, label });
  }

  return options.sort((a, b) => a.label.localeCompare(b.label));
}

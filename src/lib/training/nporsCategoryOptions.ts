/** Live SharePoint NPORS Category multi-choice values. */
export const NPORS_CATEGORY_CHOICES = [
  "N001",
  "N003",
  "N004",
  "N010",
  "N020",
  "N027",
] as const;

/**
 * Admin NPORS category dropdown — values must match SharePoint choice options.
 */
export function getNporsCategoryOptions(): Array<{
  value: string;
  label: string;
}> {
  return NPORS_CATEGORY_CHOICES.map((code) => ({
    value: code,
    label: code,
  }));
}

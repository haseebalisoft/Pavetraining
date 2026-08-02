/** Live SharePoint EUSR Category multi-choice values. */
export const EUSR_CATEGORY_CHOICES = [
  "Water Hygiene",
  "SHEA Water",
  "SHEA Gas",
  "SHEA Power",
  "SHEA Telecoms",
  "Other",
] as const;

/** Live SharePoint Card Type choice values. */
export const EUSR_CARD_TYPE_CHOICES = [
  "Physical Card",
  "Digital Card",
  "Certificate",
  "Other",
] as const;

const EUSR_CATEGORY_LABELS: Record<(typeof EUSR_CATEGORY_CHOICES)[number], string> = {
  "Water Hygiene": "National Water Hygiene",
  "SHEA Water": "SHEA Water",
  "SHEA Gas": "SHEA Gas",
  "SHEA Power": "SHEA Power",
  "SHEA Telecoms": "SHEA Telecoms",
  Other: "Other",
};

const EUSR_CARD_TYPE_LABELS: Record<(typeof EUSR_CARD_TYPE_CHOICES)[number], string> = {
  "Physical Card": "Physical",
  "Digital Card": "Virtual",
  Certificate: "Certificate",
  Other: "Other",
};

function toOptions(
  values: readonly string[],
  labels?: Record<string, string>,
) {
  return values.map((value) => ({
    value,
    label: labels?.[value] ?? value,
  }));
}

export function getEusrCategoryOptions() {
  return toOptions(EUSR_CATEGORY_CHOICES, EUSR_CATEGORY_LABELS);
}

export function getEusrCardTypeOptions() {
  return toOptions(EUSR_CARD_TYPE_CHOICES, EUSR_CARD_TYPE_LABELS);
}

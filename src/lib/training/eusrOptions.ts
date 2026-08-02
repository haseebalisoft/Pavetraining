/** Live SharePoint EUSR Category choice values. */
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

function toOptions(values: readonly string[]) {
  return values.map((value) => ({ value, label: value }));
}

export function getEusrCategoryOptions() {
  return toOptions(EUSR_CATEGORY_CHOICES);
}

export function getEusrCardTypeOptions() {
  return toOptions(EUSR_CARD_TYPE_CHOICES);
}

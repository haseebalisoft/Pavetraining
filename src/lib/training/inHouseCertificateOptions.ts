/** Live SharePoint In-House Certificate Category choice values. */
export const IN_HOUSE_CERTIFICATE_CATEGORY_CHOICES = [
  "Manual Handling",
  "Working at Height",
  "Abrasive Wheels",
  "Fire Marshal",
  "First Aid",
  "Plant Awareness",
  "Other",
] as const;

export function getInHouseCertificateCategoryOptions() {
  return IN_HOUSE_CERTIFICATE_CATEGORY_CHOICES.map((value) => ({
    value,
    label: value,
  }));
}

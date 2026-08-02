/** Live SharePoint NVQ Stage choice values. */
export const NVQ_STAGE_CHOICES = [
  "Enquiry",
  "Awaiting Details",
  "Registered",
  "Induction Booked",
  "In Progress",
  "Visit Booked",
  "Awaiting Evidence",
  "Completed",
] as const;

export function getNvqStageOptions() {
  return NVQ_STAGE_CHOICES.map((value) => ({
    value,
    label: value,
  }));
}

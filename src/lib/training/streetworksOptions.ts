/**
 * Streetworks Course dropdown — consolidated novice/reassessment into Course.
 * Values must match SharePoint NRSWA Course choice options (update SP if needed).
 */
export const STREETWORKS_COURSE_CHOICES = [
  "Operative",
  "Supervisor",
  "Operative Reassessment",
  "Supervisor Reassessment",
] as const;

/** Live SharePoint Streetworks Category multi-choice values. */
export const STREETWORKS_CATEGORY_CHOICES = [
  "LA",
  "01",
  "02",
  "03",
  "04",
  "05",
  "06",
  "07",
  "08",
  "S1",
  "S2",
  "S3",
  "S4",
  "S5",
  "S6",
  "S7",
] as const;

function toOptions(values: readonly string[], labels?: Record<string, string>) {
  return values.map((value) => ({
    value,
    label: labels?.[value] ?? value,
  }));
}

export function getStreetworksCourseOptions() {
  return toOptions(STREETWORKS_COURSE_CHOICES);
}

export function getStreetworksCategoryOptions() {
  return toOptions(STREETWORKS_CATEGORY_CHOICES);
}

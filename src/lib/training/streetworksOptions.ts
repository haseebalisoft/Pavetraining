/** Live SharePoint NRSWA Course choice values (primary admin options). */
export const STREETWORKS_COURSE_CHOICES = [
  "Operative",
  "Supervisor",
  "Refresher",
  "Unit 1",
  "Unit 2",
  "Unit 10",
  "other",
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

const COURSE_LABELS: Record<string, string> = {
  Operative: "Operative",
  Supervisor: "Supervisor",
  Refresher: "Refresher / Reassessment",
  "Unit 1": "Unit 1",
  "Unit 2": "Unit 2",
  "Unit 10": "Unit 10",
  other: "Other",
};

function toOptions(values: readonly string[], labels?: Record<string, string>) {
  return values.map((value) => ({
    value,
    label: labels?.[value] ?? value,
  }));
}

export function getStreetworksCourseOptions() {
  return toOptions(STREETWORKS_COURSE_CHOICES, COURSE_LABELS);
}

export function getStreetworksCategoryOptions() {
  return toOptions(STREETWORKS_CATEGORY_CHOICES);
}

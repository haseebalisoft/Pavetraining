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

/** Matrix column header for one EUSR category (independent training + expiry). */
export function eusrMatrixHeader(category: string): string {
  return `EUSR - ${category.trim()}`;
}

export const EUSR_MATRIX_CATEGORY_COLUMNS = EUSR_CATEGORY_CHOICES.map(
  (category) => ({
    category,
    header: eusrMatrixHeader(category),
  }),
);

export function parseEusrCategories(value: unknown): string[] {
  const tokens: string[] = [];
  if (Array.isArray(value)) {
    for (const entry of value) {
      if (typeof entry === "string" && entry.trim()) tokens.push(entry.trim());
    }
  } else if (typeof value === "string" && value.trim()) {
    tokens.push(
      ...value.split(/[;,#|]+/).map((part) => part.trim()).filter(Boolean),
    );
  }
  const seen = new Set<string>();
  const out: string[] = [];
  for (const token of tokens) {
    const key = token.toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    const known = EUSR_CATEGORY_CHOICES.find((choice) => choice.toLowerCase() === key);
    out.push(known ?? token);
  }
  return out;
}

export function eusrMatrixHeaderForCategory(
  category: string,
): string | null {
  const text = category.trim();
  if (!text) return null;
  const known = EUSR_MATRIX_CATEGORY_COLUMNS.find(
    (column) =>
      column.category.toLowerCase() === text.toLowerCase() ||
      column.header.toLowerCase() === text.toLowerCase(),
  );
  return known?.header ?? eusrMatrixHeader(text);
}

export interface EusrCategoryDateCell {
  category: string;
  header: string;
  trainingDate: string | null;
  expiry: string | null;
}

/** One matrix cell per EUSR category that has a training or expiry date. */
export function listEusrCategoryMatrixCells(
  columnValues: Record<string, string | null> | null | undefined,
  categoryTrainingDates?: Record<string, string | null> | null,
): EusrCategoryDateCell[] {
  const out: EusrCategoryDateCell[] = [];
  for (const { category, header } of EUSR_MATRIX_CATEGORY_COLUMNS) {
    const expiry = columnValues?.[header]?.trim() || null;
    const trainingDate = categoryTrainingDates?.[header]?.trim() || null;
    if (!expiry && !trainingDate) continue;
    out.push({ category, header, trainingDate, expiry });
  }
  return out;
}

/**
 * Split a stored EUSR row into one display row per category so Water Hygiene
 * and SHEA Water can keep independent training / expiry dates.
 */
export function expandEusrRecordsForDisplay<
  T extends { id: string; eusrCategory: string | null },
>(records: T[]): T[] {
  const out: T[] = [];
  for (const row of records) {
    const cats = parseEusrCategories(row.eusrCategory);
    if (cats.length <= 1) {
      out.push({
        ...row,
        eusrCategory: cats[0] ?? row.eusrCategory,
      });
      continue;
    }
    for (const cat of cats) {
      out.push({
        ...row,
        id: `${row.id}::${cat}`,
        eusrCategory: cat,
      });
    }
  }
  return out;
}

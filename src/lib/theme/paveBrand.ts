/**
 * PAVE brand tokens — single source of truth for Next.js + SPFx.
 * Keep CSS variables in globals.css / SPFx SCSS in sync with these values.
 */
export const PAVE_BRAND = {
  green: "#81CF43",
  greenHover: "#6BB536",
  greenTint: "#EEF9E4",
  charcoal: "#333333",
  charcoalLight: "#555555",
  textSecondary: "#767876",
  border: "#E5E7E5",
  bg: "#FAFAF9",
  surface: "#FFFFFF",
  amber: "#F0A828",
  amberTint: "#FDF3DC",
  red: "#E14B4B",
  redTint: "#FCE8E8",
  neutralTint: "#E5E7E5",
} as const;

/** Matrix grid layout tokens (px). */
export const MATRIX_LAYOUT = {
  cellMinWidth: 120,
  nameColWidth: 180,
  companyColWidth: 160,
  cellPaddingY: 10,
  cellPaddingX: 12,
  headerHeight: 44,
  rowMinHeight: 48,
} as const;

export type MatrixStatusKey =
  | "compliant"
  | "expiringSoon"
  | "expired"
  | "notApplicable";

export const MATRIX_STATUS = {
  compliant: {
    key: "compliant" as const,
    label: "Compliant",
    bg: PAVE_BRAND.greenTint,
    text: PAVE_BRAND.charcoal,
    accent: PAVE_BRAND.green,
  },
  expiringSoon: {
    key: "expiringSoon" as const,
    label: "Expiring soon",
    bg: PAVE_BRAND.amberTint,
    text: PAVE_BRAND.charcoal,
    accent: PAVE_BRAND.amber,
  },
  expired: {
    key: "expired" as const,
    label: "Expired",
    bg: PAVE_BRAND.redTint,
    text: PAVE_BRAND.charcoal,
    accent: PAVE_BRAND.red,
  },
  notApplicable: {
    key: "notApplicable" as const,
    label: "Not applicable",
    bg: PAVE_BRAND.neutralTint,
    text: PAVE_BRAND.textSecondary,
    accent: PAVE_BRAND.border,
  },
} as const;

/**
 * Normalize SharePoint / free-text overall status strings into matrix keys.
 * Handles Compliant vs compliant vs OK vs Valid, etc.
 */
export function normalizeMatrixStatus(
  raw: string | null | undefined,
): MatrixStatusKey {
  const value = (raw || "").trim().toLowerCase();
  if (!value) return "notApplicable";

  if (
    /^(compliant|valid|ok|good|current|pass|active)$/.test(value) ||
    value.includes("compliant") ||
    value.includes("valid")
  ) {
    return "compliant";
  }

  if (
    value.includes("expir") ||
    value.includes("urgent") ||
    value.includes("upcoming") ||
    value.includes("soon") ||
    value.includes("warning")
  ) {
    if (value.includes("expired") || value === "past") return "expired";
    return "expiringSoon";
  }

  if (
    value.includes("expired") ||
    value.includes("overdue") ||
    value.includes("lapsed")
  ) {
    return "expired";
  }

  if (
    value.includes("missing") ||
    value.includes("review") ||
    value.includes("n/a") ||
    value.includes("na") ||
    value.includes("none") ||
    value.includes("not applicable") ||
    value.includes("no record")
  ) {
    return "notApplicable";
  }

  return "notApplicable";
}

export function matrixStatusLabel(key: MatrixStatusKey): string {
  return MATRIX_STATUS[key].label;
}

/**
 * PAVE brand tokens — match Next.js globals / website (#81CF43, #333333).
 */
export const paveTheme = {
  green: "#81CF43",
  greenDark: "#4F8F22",
  greenTint: "#EEF9E4",
  charcoal: "#333333",
  charcoalDark: "#1F1F1F",
  fontHeading: '"Barlow Condensed", "Segoe UI", sans-serif',
  fontBody: 'Inter, "Segoe UI", system-ui, sans-serif',
  radiusLg: "20px",
  radiusPill: "999px",
} as const;

export type PaveTheme = typeof paveTheme;

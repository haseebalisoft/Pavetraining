/**
 * PAVE brand tokens — match Next.js globals / website (#81CF43, #333333).
 */
export const paveTheme = {
  green: "#81CF43",
  greenHover: "#6BB536",
  greenDark: "#4F8F22",
  greenTint: "#EEF9E4",
  charcoal: "#333333",
  charcoalLight: "#555555",
  charcoalMuted: "#767876",
  charcoalDark: "#1F1F1F",
  amber: "#F0A828",
  amberTint: "#FDF3DC",
  red: "#E14B4B",
  redTint: "#FCE8E8",
  bg: "#FAFAF9",
  surface: "#FFFFFF",
  border: "#E5E7E5",
  fontHeading: '"Barlow Condensed", "Segoe UI", sans-serif',
  fontBody: 'Inter, "Segoe UI", system-ui, sans-serif',
  radiusLg: "20px",
  radiusPill: "999px",
} as const;

export type PaveTheme = typeof paveTheme;

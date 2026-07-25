import styles from "./ui.module.css";

import type { StatusTone } from "@/lib/ui/status";

const TONE_CLASS: Record<StatusTone, string> = {
  neutral: styles.badgeNeutral,
  ok: styles.badgeOk,
  warn: styles.badgeWarn,
  danger: styles.badgeDanger,
  info: styles.badgeInfo,
  missing: styles.badgeMissing,
};

export function StatusBadge({
  label,
  tone = "neutral",
}: {
  label: string;
  tone?: StatusTone;
}) {
  return (
    <span className={`${styles.badge} ${TONE_CLASS[tone]}`}>{label}</span>
  );
}

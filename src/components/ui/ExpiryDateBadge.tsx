import { StatusBadge } from "@/components/ui/StatusBadge";
import { getExpiryStatus } from "@/lib/training/expiryFilters";
import { toneForExpiryStatus } from "@/lib/ui/status";
import { formatDate } from "@/lib/utils/formatDate";

import styles from "./ExpiryDateBadge.module.css";

/**
 * Colour-coded expiry cell used across customer + admin tables.
 * Status tint comes from global matrix tokens — same shape every cell.
 */
export function ExpiryDateBadge({
  date,
  showDate = true,
  fillCell = false,
}: {
  date: string | null | undefined;
  /** When false, only the status label badge is shown (e.g. summary chips). */
  showDate?: boolean;
  /** When true, fills the table cell background with status tint. */
  fillCell?: boolean;
}) {
  const status = getExpiryStatus(date);
  const fillClass =
    status.status === "valid"
      ? styles.fillCompliant
      : status.status === "expired" || status.status === "urgent"
        ? styles.fillExpired
        : status.status === "upcoming"
          ? styles.fillExpiring
          : styles.fillNa;

  return (
    <span
      className={`${styles.wrap} ${fillCell ? `${styles.fill} ${fillClass}` : ""}`}
    >
      <StatusBadge
        label={status.label}
        tone={toneForExpiryStatus(status.status)}
      />
      {showDate && status.status !== "missing" ? (
        <span className={styles.date}>{formatDate(date)}</span>
      ) : null}
    </span>
  );
}

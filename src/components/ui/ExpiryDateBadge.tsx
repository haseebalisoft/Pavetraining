import { StatusBadge } from "@/components/ui/StatusBadge";
import {
  formatDisplayDate,
  getExpiryStatus,
} from "@/lib/training/expiryFilters";
import { toneForExpiryStatus } from "@/lib/ui/status";

import styles from "./ExpiryDateBadge.module.css";

/**
 * Colour-coded expiry cell used across customer + admin tables.
 */
export function ExpiryDateBadge({
  date,
  showDate = true,
}: {
  date: string | null | undefined;
  /** When false, only the status label badge is shown (e.g. summary chips). */
  showDate?: boolean;
}) {
  const status = getExpiryStatus(date);

  return (
    <span className={styles.wrap}>
      <StatusBadge
        label={status.label}
        tone={toneForExpiryStatus(status.status)}
      />
      {showDate && status.status !== "missing" ? (
        <span className={styles.date}>{formatDisplayDate(date)}</span>
      ) : null}
    </span>
  );
}

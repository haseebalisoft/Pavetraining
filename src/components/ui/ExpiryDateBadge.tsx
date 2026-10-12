import type { ReactNode } from "react";

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
  heading,
}: {
  date: string | null | undefined;
  /** When false, only the status label badge is shown (e.g. summary chips). */
  showDate?: boolean;
  /** When true, fills the table cell background with status tint. */
  fillCell?: boolean;
  /**
   * Line shown above the status + date (e.g. a card number or category
   * title), so the thing that expires and its expiry read as one block.
   */
  heading?: ReactNode;
}) {
  const status = getExpiryStatus(date);
  const fillClass =
    status.status === "valid"
      ? styles.fillCompliant
      : status.status === "expired"
        ? styles.fillExpired
        : status.status === "urgent" || status.status === "upcoming"
          ? styles.fillExpiring
          : styles.fillNa;

  const badge = (
    <>
      <StatusBadge
        label={status.label}
        tone={toneForExpiryStatus(status.status)}
      />
      {showDate && status.status !== "missing" ? (
        <span className={styles.date}>{formatDate(date)}</span>
      ) : null}
    </>
  );

  if (heading !== undefined) {
    return (
      <span
        className={`${styles.wrap} ${styles.stacked} ${fillCell ? `${styles.fill} ${fillClass}` : ""}`}
      >
        <span className={styles.heading}>{heading}</span>
        <span className={styles.wrap}>{badge}</span>
      </span>
    );
  }

  return (
    <span
      className={`${styles.wrap} ${fillCell ? `${styles.fill} ${fillClass}` : ""}`}
    >
      {badge}
    </span>
  );
}

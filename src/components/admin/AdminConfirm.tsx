"use client";

import styles from "./admin.module.css";

interface AdminConfirmProps {
  open: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  onCancel: () => void;
  onConfirm: () => void;
}

export function AdminConfirm({
  open,
  title,
  message,
  confirmLabel = "Confirm",
  onCancel,
  onConfirm,
}: AdminConfirmProps) {
  if (!open) {
    return null;
  }

  return (
    <div className={styles.confirmBackdrop} role="alertdialog" aria-modal="true">
      <div className={styles.confirmCard}>
        <h3>{title}</h3>
        <p>{message}</p>
        <div className={styles.confirmActions}>
          <button type="button" className={styles.secondaryButton} onClick={onCancel}>
            Cancel
          </button>
          <button type="button" className={styles.primaryButton} onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

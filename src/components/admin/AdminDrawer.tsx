"use client";

import type { ReactNode } from "react";

import styles from "./admin.module.css";

interface AdminDrawerProps {
  open: boolean;
  title: string;
  onClose: () => void;
  children: ReactNode;
  footer: ReactNode;
  wide?: boolean;
}

export function AdminDrawer({
  open,
  title,
  onClose,
  children,
  footer,
  wide = false,
}: AdminDrawerProps) {
  if (!open) {
    return null;
  }

  return (
    <>
      <button
        type="button"
        className={styles.drawerBackdrop}
        aria-label="Close drawer"
        onClick={onClose}
      />
      <aside
        className={`${styles.drawer} ${wide ? styles.drawerWide : ""}`}
        role="dialog"
        aria-modal="true"
      >
        <div className={styles.drawerHeader}>
          <h2>{title}</h2>
          <button type="button" className={styles.secondaryButton} onClick={onClose}>
            Close
          </button>
        </div>
        <div className={styles.drawerBody}>{children}</div>
        <div className={styles.drawerFooter}>{footer}</div>
      </aside>
    </>
  );
}

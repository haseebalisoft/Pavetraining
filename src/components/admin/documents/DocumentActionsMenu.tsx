"use client";

import { useEffect, useId, useRef, useState } from "react";
import { AnimatePresence, motion } from "framer-motion";

import type { AdminDocumentRecord } from "@/lib/services/adminCrudService";

import styles from "./documentsBrowse.module.css";

export function DocumentActionsMenu({
  row,
  busy,
  onEditMetadata,
  onSetVisibility,
}: {
  row: AdminDocumentRecord;
  busy: boolean;
  onEditMetadata: () => void;
  onSetVisibility: (customerVisible: boolean) => void;
}) {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement>(null);
  const menuId = useId();

  useEffect(() => {
    if (!open) return;

    function onPointerDown(event: MouseEvent) {
      if (!wrapRef.current?.contains(event.target as Node)) {
        setOpen(false);
      }
    }

    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") setOpen(false);
    }

    document.addEventListener("mousedown", onPointerDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("mousedown", onPointerDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [open]);

  return (
    <div className={styles.actionsWrap} ref={wrapRef}>
      <motion.button
        type="button"
        className={styles.menuTrigger}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={menuId}
        disabled={busy}
        whileTap={{ scale: 0.96, transition: { duration: 0.15 } }}
        onClick={() => setOpen((current) => !current)}
      >
        Actions
        <span aria-hidden="true">▾</span>
      </motion.button>

      <AnimatePresence>
        {open ? (
          <motion.div
            id={menuId}
            role="menu"
            className={styles.menuPanel}
            initial={{ opacity: 0, scale: 0.94, y: -4 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.96, y: -2 }}
            transition={{ duration: 0.18, ease: "easeOut" }}
          >
            {row.previewPath ? (
              <a
                role="menuitem"
                className={styles.menuItem}
                href={row.previewPath}
                target="_blank"
                rel="noreferrer"
                onClick={() => setOpen(false)}
              >
                Preview
              </a>
            ) : null}

            {row.downloadPath ? (
              <a
                role="menuitem"
                className={styles.menuItem}
                href={row.downloadPath}
                onClick={() => setOpen(false)}
              >
                Download
              </a>
            ) : (
              <span
                role="menuitem"
                className={`${styles.menuItem} ${styles.menuItemMuted}`}
              >
                No file
              </span>
            )}

            <button
              type="button"
              role="menuitem"
              className={styles.menuItem}
              onClick={() => {
                setOpen(false);
                onEditMetadata();
              }}
            >
              Edit metadata
            </button>

            {!row.customerVisible ? (
              <button
                type="button"
                role="menuitem"
                className={styles.menuItem}
                disabled={busy}
                onClick={() => {
                  setOpen(false);
                  onSetVisibility(true);
                }}
              >
                Mark customer visible
              </button>
            ) : (
              <button
                type="button"
                role="menuitem"
                className={styles.menuItem}
                disabled={busy}
                onClick={() => {
                  setOpen(false);
                  onSetVisibility(false);
                }}
              >
                Hide from customer
              </button>
            )}
          </motion.div>
        ) : null}
      </AnimatePresence>
    </div>
  );
}

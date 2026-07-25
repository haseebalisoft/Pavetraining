"use client";

import { motion } from "framer-motion";

import styles from "./documentsBrowse.module.css";

const springHover = { type: "spring" as const, stiffness: 420, damping: 28 };

function FolderGlyph() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" fill="currentColor">
      <path d="M3.5 6.75A2.25 2.25 0 0 1 5.75 4.5h4.1c.45 0 .88.18 1.2.5l1.2 1.2c.32.32.75.5 1.2.5h5c1.24 0 2.25 1.01 2.25 2.25v8.3A2.25 2.25 0 0 1 17.45 19.5H5.75A2.25 2.25 0 0 1 3.5 17.25V6.75Z" />
    </svg>
  );
}

export function FolderCard({
  name,
  countLabel,
  onOpen,
  index,
}: {
  name: string;
  countLabel: string;
  onOpen: () => void;
  index: number;
}) {
  return (
    <motion.button
      type="button"
      className={styles.folderCard}
      onClick={onOpen}
      initial={{ opacity: 0, y: 12 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.22, delay: index * 0.05, ease: "easeOut" }}
      whileHover={{
        scale: 1.02,
        boxShadow: "0 10px 28px rgba(74, 74, 74, 0.12)",
        transition: springHover,
      }}
      whileTap={{ scale: 0.98, transition: { duration: 0.15 } }}
    >
      <span className={styles.folderIcon}>
        <FolderGlyph />
      </span>
      <p className={styles.folderName}>{name}</p>
      <p className={styles.folderCount}>{countLabel}</p>
    </motion.button>
  );
}

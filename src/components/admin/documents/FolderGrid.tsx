"use client";

import { FolderCard } from "./FolderCard";
import styles from "./documentsBrowse.module.css";

export interface FolderItem {
  key: string;
  name: string;
  count: number;
  countNoun?: string;
}

export function FolderGrid({
  items,
  onOpen,
  countNoun = "items",
}: {
  items: FolderItem[];
  onOpen: (key: string) => void;
  countNoun?: string;
}) {
  return (
    <div className={styles.folderGrid}>
      {items.map((item, index) => (
        <FolderCard
          key={item.key}
          name={item.name}
          countLabel={`${item.count} ${item.countNoun ?? countNoun}`}
          index={index}
          onOpen={() => onOpen(item.key)}
        />
      ))}
    </div>
  );
}

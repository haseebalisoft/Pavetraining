"use client";

import { useCallback } from "react";

import styles from "./stickyLeadColumns.module.css";

/** Highest pin count the stylesheet has rules for. */
export const MAX_STICKY_LEAD_COLUMNS = 4;

/**
 * Pins the first `count` columns of a wide table while it scrolls sideways.
 *
 * Fixed rem offsets drift from the real column widths (long names, padding,
 * table-layout), which let a pinned column slide over the one before it. This
 * measures the rendered header cells instead and publishes the running left
 * offsets as `--sticky-left-1 … --sticky-left-N` on the table, re-measuring
 * whenever a pinned column changes size.
 *
 * Pass the returned callback as the `<table>` ref, and spread
 * `stickyLeadTableProps(count)` onto the same table.
 */
export function useStickyLeadColumns(
  count: number,
): (table: HTMLTableElement | null) => (() => void) | void {
  const pinned = clampCount(count);

  const ref = useCallback(
    (table: HTMLTableElement | null) => {
      if (!table || pinned === 0) return;

      const measure = () => {
        const headerRow = table.tHead?.rows[0];
        if (!headerRow) return;
        let left = 0;
        for (let index = 0; index < pinned; index += 1) {
          const cell = headerRow.cells[index];
          if (!cell) break;
          table.style.setProperty(`--sticky-left-${index + 1}`, `${left}px`);
          left += cell.getBoundingClientRect().width;
        }
      };

      measure();
      if (typeof ResizeObserver === "undefined") return;
      const observer = new ResizeObserver(measure);
      observer.observe(table);
      const headerRow = table.tHead?.rows[0];
      for (let index = 0; index < pinned; index += 1) {
        const cell = headerRow?.cells[index];
        if (cell) observer.observe(cell);
      }
      return () => observer.disconnect();
    },
    [pinned],
  );

  return ref;
}

/** Class + data attribute that switch on the pinned-column styles. */
export function stickyLeadTableProps(count: number): {
  className: string;
  "data-sticky-count": number | undefined;
} {
  const pinned = clampCount(count);
  return {
    className: pinned > 0 ? styles.stickyLead : "",
    "data-sticky-count": pinned > 0 ? pinned : undefined,
  };
}

function clampCount(count: number): number {
  return Math.max(0, Math.min(MAX_STICKY_LEAD_COLUMNS, Math.floor(count)));
}

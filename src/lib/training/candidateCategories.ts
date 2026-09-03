/**
 * Aggregates a candidate's training records from every register plus the
 * Training Matrix's own N-code/CSCS columns into one sorted list — the single
 * source both the customer and admin candidate profile render, so the two
 * views can never show different data for the same candidate.
 */
import { CLIENT_MATRIX_CATEGORY_COLUMNS } from "@/lib/services/bulkUpload/clientTemplateHeaders";
import {
  MATRIX_EDITOR_DATE_HEADERS,
  matrixEditorSectionTitle,
  matrixStoredExpiryKey,
} from "@/lib/training/matrixEditorFields";
import { parseEusrCategories } from "@/lib/training/eusrOptions";
import { getExpiryStatus, type ExpiryStatus } from "@/lib/training/expiryFilters";
import type {
  CustomerEusrRecord,
  CustomerInHouseRecord,
  CustomerMatrixRecord,
  CustomerNporsRecord,
  CustomerNvqRecord,
  CustomerStreetworksRecord,
} from "@/types/models";

export type CandidateCategorySource =
  | "NPORS"
  | "EUSR"
  | "Streetworks / NRSWA"
  | "In-House"
  | "NVQ"
  | "Training Matrix";

export interface CandidateCategoryRow {
  id: string;
  category: string;
  source: CandidateCategorySource;
  trainingDate: string | null;
  expiryDate: string | null;
  /** Free-text outcome — "Pass"/"Fail" for registers, "Active"/"Completed" for NVQ. */
  outcomeLabel: string | null;
  expiryStatus: ExpiryStatus;
}

const SORT_RANK: Record<ExpiryStatus["status"], number> = {
  expired: 0,
  urgent: 1,
  upcoming: 1,
  valid: 2,
  missing: 3,
};

const EUSR_ROLLUP_HEADER = "EUSR Expiry";
const NRSWA_ROLLUP_HEADER = "NRSWA Expiry";

/** Named matrix fields used when `columnValues` is not on the customer row. */
const NAMED_EXPIRY: Record<
  string,
  (row: CustomerMatrixRecord) => string | null | undefined
> = {
  "CSCS Expiry": (row) => row.cscsExpiry,
  "NRSWA Expiry": (row) => row.swqrExpiry,
  "EUSR Expiry": (row) => row.eusrExpiry,
  "N001 - Ind FLT": (row) => row.n001Expiry,
  "N003 - Reach Lift Truck": (row) => row.n003Expiry,
  "N004 - Lorry Mounted Lift Truck": (row) => row.n004Expiry,
  "N010 - Telescopic Handler": (row) => row.n010Expiry,
  "N020 - Tiltrotator System": (row) => row.n020Expiry,
  "N021 - Suction Excavator": (row) => row.n021Expiry,
  "N027 - Excavation Marshal - Banksperson": (row) => row.n027Expiry,
  "N100 - Exc Crane": (row) => row.n100Expiry,
  "N031 - Asbestos Awareness": (row) => row.inHouseExpiry,
};

function hasAnyDate(
  ...values: Array<string | null | undefined>
): boolean {
  return values.some((value) => Boolean(value?.trim()));
}

function toRow(
  id: string,
  category: string | null | undefined,
  fallbackCategory: string,
  source: CandidateCategorySource,
  trainingDate: string | null,
  expiryDate: string | null,
  outcomeLabel: string | null,
): CandidateCategoryRow {
  return {
    id,
    category: category?.trim() || fallbackCategory,
    source,
    trainingDate,
    expiryDate,
    outcomeLabel,
    expiryStatus: getExpiryStatus(expiryDate),
  };
}

function coverKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

function addCovered(covered: Set<string>, ...values: Array<string | null | undefined>) {
  for (const value of values) {
    const key = coverKey(value);
    if (key) covered.add(key);
  }
}

function isCovered(covered: Set<string>, ...values: Array<string | null | undefined>) {
  return values.some((value) => covered.has(coverKey(value)));
}

/**
 * True when a dated NPORS register row already represents this matrix header.
 * Blank-date NPORS rows must not cover — otherwise matrix dates (including
 * expired) disappear from the profile with nothing to replace them.
 */
function nporsCoversHeader(
  records: CustomerNporsRecord[],
  header: string,
): boolean {
  const column = CLIENT_MATRIX_CATEGORY_COLUMNS.find(
    (entry) => entry.header.toLowerCase() === header.toLowerCase(),
  );
  for (const row of records) {
    if (!hasAnyDate(row.trainingDate, row.expiry)) continue;
    const category = row.nporsCategory?.trim().toLowerCase() ?? "";
    if (!category) continue;
    if (category === header.toLowerCase()) return true;
    if (!column) continue;
    if (
      category === column.code.toLowerCase() ||
      category === column.header.toLowerCase() ||
      category.includes(column.code.toLowerCase())
    ) {
      return true;
    }
  }
  return false;
}

function matrixExpiry(
  row: CustomerMatrixRecord,
  header: string,
): string | null {
  const fromColumns = row.columnValues?.[header]?.trim() || null;
  if (fromColumns) return fromColumns;
  const stored =
    row.categoryTrainingDates?.[matrixStoredExpiryKey(header)]?.trim() || null;
  if (stored) return stored;
  return NAMED_EXPIRY[header]?.(row)?.trim() || null;
}

function matrixTraining(
  row: CustomerMatrixRecord,
  header: string,
): string | null {
  const dates = row.categoryTrainingDates;
  if (!dates) return null;
  return (
    dates[header]?.trim() ||
    dates[`${header} Training Date`]?.trim() ||
    null
  );
}

export function buildCandidateCategoryRows(input: {
  nporsRecords?: CustomerNporsRecord[];
  eusrRecords?: CustomerEusrRecord[];
  streetworksRecords?: CustomerStreetworksRecord[];
  inHouseRecords?: CustomerInHouseRecord[];
  nvqRecords?: CustomerNvqRecord[];
  matrixRow?: CustomerMatrixRecord | null;
}): CandidateCategoryRow[] {
  const rows: CandidateCategoryRow[] = [];
  const covered = new Set<string>();

  for (const row of input.nporsRecords ?? []) {
    if (!hasAnyDate(row.trainingDate, row.expiry)) continue;
    const next = toRow(
      `npors:${row.id}`,
      row.nporsCategory,
      "NPORS",
      "NPORS",
      row.trainingDate,
      row.expiry,
      row.outcome,
    );
    rows.push(next);
    addCovered(covered, next.category);
    const column = CLIENT_MATRIX_CATEGORY_COLUMNS.find(
      (entry) =>
        entry.header.toLowerCase() === next.category.toLowerCase() ||
        entry.code.toLowerCase() === next.category.toLowerCase() ||
        next.category.toLowerCase().includes(entry.code.toLowerCase()),
    );
    if (column) addCovered(covered, column.header, column.code);
  }

  for (const row of input.eusrRecords ?? []) {
    if (!hasAnyDate(row.trainingDate, row.expiry)) continue;
    const categories = parseEusrCategories(row.eusrCategory);
    const labels = categories.length === 0 ? [row.eusrCategory] : categories;
    for (const category of labels) {
      const next = toRow(
        `eusr:${row.id}:${category ?? "eusr"}`,
        category,
        "EUSR",
        "EUSR",
        row.trainingDate,
        row.expiry,
        row.outcome,
      );
      rows.push(next);
      addCovered(covered, next.category, `EUSR - ${next.category}`, EUSR_ROLLUP_HEADER);
    }
  }

  for (const row of input.streetworksRecords ?? []) {
    if (!hasAnyDate(row.trainingDate, row.expiry)) continue;
    const next = toRow(
      `streetworks:${row.id}`,
      row.streetworksCategory ?? row.course,
      "Streetworks / NRSWA",
      "Streetworks / NRSWA",
      row.trainingDate,
      row.expiry,
      row.outcome,
    );
    rows.push(next);
    addCovered(covered, next.category, NRSWA_ROLLUP_HEADER);
  }

  for (const row of input.inHouseRecords ?? []) {
    if (!hasAnyDate(row.trainingDate, row.expiry)) continue;
    const next = toRow(
      `inhouse:${row.id}`,
      row.course,
      "In-House",
      "In-House",
      row.trainingDate,
      row.expiry,
      row.outcome,
    );
    rows.push(next);
    addCovered(covered, next.category);
    if (/asbestos/i.test(next.category)) {
      addCovered(covered, "N031 - Asbestos Awareness");
    }
  }

  for (const row of input.nvqRecords ?? []) {
    if (
      !hasAnyDate(row.dateRegistered, row.inductionDate, row.completedDate)
    ) {
      continue;
    }
    const next = toRow(
      `nvq:${row.id}`,
      row.nvqTitle,
      "NVQ",
      "NVQ",
      row.dateRegistered ?? row.inductionDate ?? row.completedDate,
      // NVQ has no expiry concept — status/outcome carries the record instead.
      null,
      row.status,
    );
    rows.push(next);
    addCovered(covered, next.category);
  }

  if (input.matrixRow) {
    for (const cell of input.matrixRow.eusrCategoryRows ?? []) {
      if (!hasAnyDate(cell.trainingDate, cell.expiry)) continue;
      if (isCovered(covered, cell.category, `EUSR - ${cell.category}`)) continue;
      const next = toRow(
        `matrix:eusr:${cell.category}`,
        cell.category,
        "EUSR",
        "Training Matrix",
        cell.trainingDate,
        cell.expiry,
        null,
      );
      rows.push(next);
      addCovered(
        covered,
        next.category,
        `EUSR - ${next.category}`,
        EUSR_ROLLUP_HEADER,
      );
    }

    for (const header of MATRIX_EDITOR_DATE_HEADERS) {
      const trainingDate = matrixTraining(input.matrixRow, header);
      const expiryDate = matrixExpiry(input.matrixRow, header);
      if (!hasAnyDate(trainingDate, expiryDate)) continue;

      const label = matrixEditorSectionTitle(header);
      if (
        isCovered(covered, header, label) ||
        nporsCoversHeader(input.nporsRecords ?? [], header)
      ) {
        continue;
      }

      const next = toRow(
        `matrix:${header}`,
        label,
        label,
        "Training Matrix",
        trainingDate,
        expiryDate,
        null,
      );
      rows.push(next);
      addCovered(covered, header, label, next.category);
    }
  }

  return rows.sort((a, b) => {
    const rankDiff =
      SORT_RANK[a.expiryStatus.status] - SORT_RANK[b.expiryStatus.status];
    if (rankDiff !== 0) return rankDiff;

    const daysA = a.expiryStatus.daysUntilExpiry;
    const daysB = b.expiryStatus.daysUntilExpiry;
    if (daysA !== null && daysB !== null && daysA !== daysB) {
      return daysA - daysB;
    }

    return a.category.localeCompare(b.category);
  });
}

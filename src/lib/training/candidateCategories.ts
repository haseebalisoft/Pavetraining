/**
 * Aggregates a candidate's training records from every register plus the
 * Training Matrix's own N-code/CSCS columns into one sorted list — the single
 * source both the customer and admin candidate profile render, so the two
 * views can never show different data for the same candidate.
 */
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

/**
 * Matrix columns with no dedicated register elsewhere in this app — CSCS has
 * no register list at all, and the N-code plant tickets are held per-piece-
 * of-equipment on the matrix, not one-row-per-code on the NPORS register.
 * Anything already covered by a register (NPORS/EUSR/Streetworks/In-House) is
 * deliberately left out here so the same category never appears twice.
 */
const MATRIX_ONLY_CATEGORIES: ReadonlyArray<{
  category: string;
  pick: (row: CustomerMatrixRecord) => string | null;
}> = [
  { category: "CSCS", pick: (row) => row.cscsExpiry },
  { category: "N001 - Ind FLT", pick: (row) => row.n001Expiry },
  { category: "N003 - Reach Lift Truck", pick: (row) => row.n003Expiry },
  {
    category: "N004 - Lorry Mounted Lift Truck",
    pick: (row) => row.n004Expiry,
  },
  { category: "N010 - Telescopic Handler", pick: (row) => row.n010Expiry },
  { category: "N020 - Tiltrotator System", pick: (row) => row.n020Expiry },
  { category: "N021 - Suction Excavator", pick: (row) => row.n021Expiry },
  {
    category: "N027 - Excavation Marshal - Banksperson",
    pick: (row) => row.n027Expiry,
  },
  { category: "N100 - Exc Crane", pick: (row) => row.n100Expiry },
  { category: "N031 - Asbestos Awareness", pick: (row) => row.inHouseExpiry },
];

export function buildCandidateCategoryRows(input: {
  nporsRecords?: CustomerNporsRecord[];
  eusrRecords?: CustomerEusrRecord[];
  streetworksRecords?: CustomerStreetworksRecord[];
  inHouseRecords?: CustomerInHouseRecord[];
  nvqRecords?: CustomerNvqRecord[];
  matrixRow?: CustomerMatrixRecord | null;
}): CandidateCategoryRow[] {
  const rows: CandidateCategoryRow[] = [];

  for (const row of input.nporsRecords ?? []) {
    rows.push(
      toRow(
        `npors:${row.id}`,
        row.nporsCategory,
        "NPORS",
        "NPORS",
        row.trainingDate,
        row.expiry,
        row.outcome,
      ),
    );
  }

  for (const row of input.eusrRecords ?? []) {
    rows.push(
      toRow(
        `eusr:${row.id}`,
        row.eusrCategory,
        "EUSR",
        "EUSR",
        row.trainingDate,
        row.expiry,
        row.outcome,
      ),
    );
  }

  for (const row of input.streetworksRecords ?? []) {
    rows.push(
      toRow(
        `streetworks:${row.id}`,
        row.streetworksCategory ?? row.course,
        "Streetworks / NRSWA",
        "Streetworks / NRSWA",
        row.trainingDate,
        row.expiry,
        row.outcome,
      ),
    );
  }

  for (const row of input.inHouseRecords ?? []) {
    rows.push(
      toRow(
        `inhouse:${row.id}`,
        row.course,
        "In-House",
        "In-House",
        row.trainingDate,
        row.expiry,
        row.outcome,
      ),
    );
  }

  for (const row of input.nvqRecords ?? []) {
    rows.push(
      toRow(
        `nvq:${row.id}`,
        row.nvqTitle,
        "NVQ",
        "NVQ",
        row.dateRegistered ?? row.inductionDate,
        // NVQ has no expiry concept — status/outcome carries the record instead.
        null,
        row.status,
      ),
    );
  }

  if (input.matrixRow) {
    for (const { category, pick } of MATRIX_ONLY_CATEGORIES) {
      const expiry = pick(input.matrixRow);
      if (!expiry?.trim()) continue;
      rows.push(
        toRow(
          `matrix:${category}`,
          category,
          category,
          "Training Matrix",
          null,
          expiry,
          null,
        ),
      );
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

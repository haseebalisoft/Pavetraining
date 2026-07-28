import "server-only";

import { listAdminMatrix } from "@/lib/services/adminCrudService";
import {
  filterRowsByCandidateAccess,
  getAllowedCandidateNames,
  getAllowedWorkforceForCustomer,
} from "@/lib/services/customerAccessService";
import {
  getCustomerDocumentRecords,
  getCustomerEventRecords,
  getCustomerNvqRecords,
  getCustomerOfferRecords,
} from "@/lib/services/customerPortalService";
import {
  getCustomerEusrRecords,
  getCustomerInHouseRecords,
  getCustomerNporsRecords,
  getCustomerStreetworksRecords,
} from "@/lib/services/customerTrainingRecordsService";
import {
  earliestExpiryDate,
  getExpiryStatus,
} from "@/lib/training/expiryFilters";
import type {
  CustomerContext,
  CustomerMatrixRecord,
  DashboardStats,
  WorkforceCandidate,
} from "@/types/models";

const NPORS_CATEGORY_COLUMNS = [
  { code: "N001", key: "n001Expiry" as const },
  { code: "N003", key: "n003Expiry" as const },
  { code: "N004", key: "n004Expiry" as const },
  { code: "N010", key: "n010Expiry" as const },
  { code: "N020", key: "n020Expiry" as const },
  { code: "N021", key: "n021Expiry" as const },
  { code: "N027", key: "n027Expiry" as const },
  { code: "N100", key: "n100Expiry" as const },
] as const;

type MatrixSourceRow = Awaited<ReturnType<typeof listAdminMatrix>>[number];

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

function emptyMatrixDates(): Pick<
  CustomerMatrixRecord,
  | "n001Expiry"
  | "n003Expiry"
  | "n004Expiry"
  | "n010Expiry"
  | "n020Expiry"
  | "n021Expiry"
  | "n027Expiry"
  | "n100Expiry"
> {
  return {
    n001Expiry: null,
    n003Expiry: null,
    n004Expiry: null,
    n010Expiry: null,
    n020Expiry: null,
    n021Expiry: null,
    n027Expiry: null,
    n100Expiry: null,
  };
}

function nporsCategoriesFromMatrix(matrix: MatrixSourceRow | null): string[] {
  if (!matrix) return [];
  const codes: string[] = [];
  for (const column of NPORS_CATEGORY_COLUMNS) {
    if (matrix[column.key]?.trim()) {
      codes.push(column.code);
    }
  }
  return codes;
}

function buildEnrichedRow(input: {
  candidate: WorkforceCandidate | null;
  matrix: MatrixSourceRow | null;
  nporsCategories: string[];
  nporsExpiry: string | null;
  nporsNumber?: string | null;
  swqrExpiry: string | null;
  eusrExpiry: string | null;
  inHouseExpiry: string | null;
  inHouseCourse?: string | null;
}): CustomerMatrixRecord {
  const { candidate, matrix } = input;
  const candidateName =
    candidate?.candidateName ?? matrix?.candidateName ?? "Unknown";
  const dates = matrix
    ? {
        n001Expiry:
          matrix.columnValues?.["N001 - Ind FLT"] ?? matrix.n001Expiry,
        n003Expiry:
          matrix.columnValues?.["N003 - Reach Lift Truck"] ?? matrix.n003Expiry,
        n004Expiry:
          matrix.columnValues?.["N004 - Lorry Mounted Lift Truck"] ??
          matrix.n004Expiry,
        n010Expiry:
          matrix.columnValues?.["N010 - Telescopic Handler"] ??
          matrix.n010Expiry,
        n020Expiry:
          matrix.columnValues?.["N020 - Tiltrotator System"] ??
          matrix.n020Expiry,
        n021Expiry:
          matrix.columnValues?.["N021 - Suction Excavator"] ?? matrix.n021Expiry,
        n027Expiry:
          matrix.columnValues?.["N027 - Excavation Marshal - Banksperson"] ??
          matrix.n027Expiry,
        n100Expiry:
          matrix.columnValues?.["N100 - Exc Crane"] ?? matrix.n100Expiry,
      }
    : emptyMatrixDates();

  const cscsExpiry =
    matrix?.columnValues?.["CSCS Expiry"] ?? candidate?.cscsExpiry ?? null;
  const swqrFromMatrix = matrix?.columnValues?.["NRSWA Expiry"] ?? null;
  const eusrFromMatrix = matrix?.columnValues?.["EUSR Expiry"] ?? null;
  const nextExpiryDate =
    earliestExpiryDate([
      matrix?.nextExpiryDate,
      input.nporsExpiry,
      cscsExpiry,
      input.swqrExpiry ?? swqrFromMatrix,
      input.eusrExpiry ?? eusrFromMatrix,
      input.inHouseExpiry,
      ...Object.values(dates),
    ]) ?? null;

  const expiryStatus = getExpiryStatus(nextExpiryDate);
  const needsReview =
    Boolean(matrix?.needsReview) || expiryStatus.status === "missing";
  const overallStatus =
    matrix?.overallStatus?.trim() ||
    (needsReview ? "Records to Review" : expiryStatus.label);

  return {
    id: candidate?.id ?? matrix?.id ?? candidateName,
    candidateId: candidate?.id ?? null,
    candidateName,
    companyName:
      candidate?.companyName ??
      matrix?.companyName ??
      null,
    dateOfBirth:
      candidate?.dateOfBirth ??
      matrix?.columnValues?.DOB ??
      matrix?.dateOfBirth ??
      null,
    department: candidate?.department ?? matrix?.department ?? null,
    trainingManager: candidate?.trainingManager ?? null,
    supervisor: candidate?.supervisor ?? null,
    overallStatus,
    needsReview,
    nextExpiryDate,
    nporsCategories: input.nporsCategories.length
      ? input.nporsCategories.join(", ")
      : null,
    nporsExpiry: input.nporsExpiry,
    nporsNumber: input.nporsNumber ?? candidate?.nporsNumbers ?? null,
    cscsNumber: candidate?.cscsNumber ?? null,
    cscsExpiry,
    swqrNumber: candidate?.swqrNumber ?? null,
    swqrExpiry: input.swqrExpiry ?? swqrFromMatrix,
    eusrNumber: candidate?.eusrNumber ?? null,
    eusrExpiry: input.eusrExpiry ?? eusrFromMatrix,
    inHouseCourse: input.inHouseCourse ?? null,
    inHouseExpiry: input.inHouseExpiry,
    ...dates,
  };
}

/**
 * Customer-facing matrix rows for a company.
 * Combines Workforce + Training Matrix + register expiries without changing
 * SharePoint list schemas. Applies Supervisor / Candidate access when context
 * is provided.
 */
export async function getCustomerMatrixRecords(
  companyName: string,
  context?: CustomerContext,
): Promise<CustomerMatrixRecord[]> {
  const companyId = context?.companyId;

  const [workforce, matrixRows, npors, eusr, streetworks, inHouse] =
    await Promise.all([
      context
        ? getAllowedWorkforceForCustomer(context)
        : Promise.resolve([] as WorkforceCandidate[]),
      listAdminMatrix(companyName),
      companyId
        ? getCustomerNporsRecords(companyId, context)
        : Promise.resolve([]),
      companyId
        ? getCustomerEusrRecords(companyId, context)
        : Promise.resolve([]),
      companyId
        ? getCustomerStreetworksRecords(companyId, context)
        : Promise.resolve([]),
      companyId
        ? getCustomerInHouseRecords(companyId, context)
        : Promise.resolve([]),
    ]);

  // When called without context (legacy), still return matrix-only rows.
  if (!context) {
    return matrixRows.map((row) =>
      buildEnrichedRow({
        candidate: null,
        matrix: row,
        nporsCategories: nporsCategoriesFromMatrix(row),
        nporsExpiry: earliestExpiryDate([
          row.nextExpiryDate,
          row.n001Expiry,
          row.n003Expiry,
          row.n004Expiry,
          row.n010Expiry,
          row.n020Expiry,
          row.n021Expiry,
          row.n027Expiry,
          row.n100Expiry,
        ]),
        swqrExpiry: null,
        eusrExpiry: null,
        inHouseExpiry: null,
      }),
    );
  }

  const matrixByName = new Map<string, MatrixSourceRow>();
  for (const row of matrixRows) {
    const key = nameKey(row.candidateName);
    if (key && !matrixByName.has(key)) {
      matrixByName.set(key, row);
    }
  }

  const nporsCatsByName = new Map<string, Set<string>>();
  const nporsExpiryByName = new Map<string, string[]>();
  const nporsNumberByName = new Map<string, string>();
  for (const row of npors) {
    const key = nameKey(row.candidateName);
    if (!key) continue;
    if (row.nporsCategory?.trim()) {
      const set = nporsCatsByName.get(key) ?? new Set<string>();
      set.add(row.nporsCategory.trim());
      nporsCatsByName.set(key, set);
    }
    if (row.nporsNumber?.trim() && !nporsNumberByName.has(key)) {
      nporsNumberByName.set(key, row.nporsNumber.trim());
    }
    if (row.expiry?.trim()) {
      const list = nporsExpiryByName.get(key) ?? [];
      list.push(row.expiry);
      nporsExpiryByName.set(key, list);
    }
  }

  const eusrExpiryByName = new Map<string, string[]>();
  for (const row of eusr) {
    const key = nameKey(row.candidateName);
    if (!key || !row.expiry?.trim()) continue;
    const list = eusrExpiryByName.get(key) ?? [];
    list.push(row.expiry);
    eusrExpiryByName.set(key, list);
  }

  const swqrExpiryByName = new Map<string, string[]>();
  for (const row of streetworks) {
    const key = nameKey(row.candidateName);
    if (!key || !row.expiry?.trim()) continue;
    const list = swqrExpiryByName.get(key) ?? [];
    list.push(row.expiry);
    swqrExpiryByName.set(key, list);
  }

  const inHouseExpiryByName = new Map<string, string[]>();
  const inHouseCourseByName = new Map<string, string>();
  for (const row of inHouse) {
    const key = nameKey(row.candidateName);
    if (!key) continue;
    if (row.course?.trim() && !inHouseCourseByName.has(key)) {
      inHouseCourseByName.set(key, row.course.trim());
    }
    if (!row.expiry?.trim()) continue;
    const list = inHouseExpiryByName.get(key) ?? [];
    list.push(row.expiry);
    inHouseExpiryByName.set(key, list);
  }

  const usedMatrix = new Set<string>();
  const rows: CustomerMatrixRecord[] = [];

  for (const candidate of workforce) {
    const key = nameKey(candidate.candidateName);
    const matrix = matrixByName.get(key) ?? null;
    if (matrix) usedMatrix.add(key);

    const matrixCats = nporsCategoriesFromMatrix(matrix);
    const registerCats = Array.from(nporsCatsByName.get(key) ?? []);
    const categories = Array.from(new Set([...matrixCats, ...registerCats]));

    const nporsExpiry = earliestExpiryDate([
      ...(nporsExpiryByName.get(key) ?? []),
      matrix?.n001Expiry,
      matrix?.n003Expiry,
      matrix?.n004Expiry,
      matrix?.n010Expiry,
      matrix?.n020Expiry,
      matrix?.n021Expiry,
      matrix?.n027Expiry,
      matrix?.n100Expiry,
    ]);

    rows.push(
      buildEnrichedRow({
        candidate,
        matrix,
        nporsCategories: categories,
        nporsExpiry,
        nporsNumber: nporsNumberByName.get(key) ?? candidate.nporsNumbers,
        swqrExpiry: earliestExpiryDate([
          candidate.swqrExpiry,
          ...(swqrExpiryByName.get(key) ?? []),
        ]),
        eusrExpiry: earliestExpiryDate([
          candidate.eusrExpiry,
          ...(eusrExpiryByName.get(key) ?? []),
        ]),
        inHouseExpiry: earliestExpiryDate(inHouseExpiryByName.get(key) ?? []),
        inHouseCourse: inHouseCourseByName.get(key) ?? null,
      }),
    );
  }

  // Matrix-only people still in scope (rare: matrix row without workforce match).
  for (const [key, matrix] of matrixByName) {
    if (usedMatrix.has(key)) continue;
    rows.push(
      buildEnrichedRow({
        candidate: null,
        matrix,
        nporsCategories: [
          ...nporsCategoriesFromMatrix(matrix),
          ...Array.from(nporsCatsByName.get(key) ?? []),
        ],
        nporsExpiry: earliestExpiryDate([
          ...(nporsExpiryByName.get(key) ?? []),
          matrix.n001Expiry,
          matrix.n003Expiry,
          matrix.n004Expiry,
          matrix.n010Expiry,
          matrix.n020Expiry,
          matrix.n021Expiry,
          matrix.n027Expiry,
          matrix.n100Expiry,
        ]),
        nporsNumber: nporsNumberByName.get(key) ?? null,
        swqrExpiry: earliestExpiryDate(swqrExpiryByName.get(key) ?? []),
        eusrExpiry: earliestExpiryDate(eusrExpiryByName.get(key) ?? []),
        inHouseExpiry: earliestExpiryDate(inHouseExpiryByName.get(key) ?? []),
        inHouseCourse: inHouseCourseByName.get(key) ?? null,
      }),
    );
  }

  const allowedNames = await getAllowedCandidateNames(context);
  return filterRowsByCandidateAccess(rows, allowedNames, context).sort((a, b) =>
    a.candidateName.localeCompare(b.candidateName),
  );
}

/**
 * Dashboard stats — only loads lists needed for the cards.
 * Training-register counts are deferred (open Training Records for detail)
 * so first paint does not wait on NPORS/EUSR/Streetworks/In-House.
 */
export async function getCustomerDashboard(
  context: CustomerContext,
): Promise<DashboardStats> {
  const [workforce, matrix, documents, events, offers, nvq] = await Promise.all([
    getAllowedWorkforceForCustomer(context),
    getCustomerMatrixRecords(context.companyName, context),
    getCustomerDocumentRecords(context.companyId, false, context),
    getCustomerEventRecords(context.companyId),
    getCustomerOfferRecords(context.companyId),
    getCustomerNvqRecords(context.companyId, context),
  ]);

  let needsReviewCount = 0;
  let expiringSoonCount = 0;
  let upcomingExpiryCount = 0;
  let expiredCount = 0;

  for (const row of matrix) {
    if (row.needsReview) {
      needsReviewCount += 1;
    }
    const status = getExpiryStatus(row.nextExpiryDate).status;
    if (status === "expired") {
      expiredCount += 1;
    } else if (status === "urgent") {
      expiringSoonCount += 1;
    } else if (status === "upcoming") {
      upcomingExpiryCount += 1;
    }
  }

  const now = Date.now();
  const upcomingEventsCount = events.filter((event) => {
    if (!event.eventDate) return false;
    const time = new Date(event.eventDate).getTime();
    return !Number.isNaN(time) && time >= now;
  }).length;

  return {
    workforceCount: workforce.length,
    trainingMatrixCount: matrix.length,
    needsReviewCount,
    expiringSoonCount,
    upcomingExpiryCount,
    expiredCount,
    documentsCount: documents.length,
    upcomingEventsCount,
    activeOffersCount: offers.length,
    nporsCount: 0,
    eusrCount: 0,
    streetworksCount: 0,
    inHouseCount: 0,
    nvqCount: nvq.length,
  };
}

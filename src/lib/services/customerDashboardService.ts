import "server-only";

import { listAdminMatrix } from "@/lib/services/adminCrudService";
import {
  earliestDateFromColumns,
  listTrainingMatrixExampleRows,
} from "@/lib/services/bulkUpload/trainingMatrixExampleService";
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
import {
  listEusrCategoryMatrixCells,
  parseEusrCategories,
} from "@/lib/training/eusrOptions";
import type {
  CustomerContext,
  CustomerDocumentRecord,
  CustomerEusrCategoryCell,
  CustomerEusrRecord,
  CustomerEventRecord,
  CustomerMatrixRecord,
  CustomerOfferRecord,
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

type MatrixSourceRow = {
  id: string;
  candidateName: string;
  companyName: string | null;
  department: string | null;
  dateOfBirth: string | null;
  overallStatus: string | null;
  needsReview: boolean;
  nextExpiryDate: string | null;
  n001Expiry: string | null;
  n003Expiry: string | null;
  n004Expiry: string | null;
  n010Expiry: string | null;
  n020Expiry: string | null;
  n021Expiry: string | null;
  n027Expiry: string | null;
  n100Expiry: string | null;
  n031Expiry?: string | null;
  columnValues: Record<string, string | null>;
  categoryTrainingDates?: Record<string, string | null>;
};

function exampleToMatrixSource(
  example: Awaited<ReturnType<typeof listTrainingMatrixExampleRows>>[number],
  companyName: string | null,
): MatrixSourceRow {
  const columnValues = { ...example.columnValues };
  const nextExpiryDate =
    example.nextExpiryDate ?? earliestDateFromColumns(columnValues);
  return {
    id: `example:${example.id}`,
    candidateName: example.candidateName,
    companyName,
    department: null,
    dateOfBirth: example.dateOfBirth,
    overallStatus: null,
    needsReview: !nextExpiryDate,
    nextExpiryDate,
    n001Expiry: columnValues["N001 - Ind FLT"] ?? null,
    n003Expiry: columnValues["N003 - Reach Lift Truck"] ?? null,
    n004Expiry: columnValues["N004 - Lorry Mounted Lift Truck"] ?? null,
    n010Expiry: columnValues["N010 - Telescopic Handler"] ?? null,
    n020Expiry: columnValues["N020 - Tiltrotator System"] ?? null,
    n021Expiry: columnValues["N021 - Suction Excavator"] ?? null,
    n027Expiry:
      columnValues["N027 - Excavation Marshal - Banksperson"] ?? null,
    n100Expiry: columnValues["N100 - Exc Crane"] ?? null,
    n031Expiry: columnValues["N031 - Asbestos Awareness"] ?? null,
    columnValues,
    categoryTrainingDates: example.categoryTrainingDates,
  };
}

function nameKey(value: string | null | undefined): string {
  // Collapses internal runs of whitespace so "John  Smith" and "John Smith"
  // resolve to the same candidate (matches candidateNameKey in the sync core).
  return (value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
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
    const fromTyped = matrix[column.key]?.trim();
    const fromColumns = Object.entries(matrix.columnValues ?? {}).find(([header]) =>
      header.toUpperCase().startsWith(`${column.code} `) ||
      header.toUpperCase().startsWith(`${column.code} -`),
    )?.[1];
    if (fromTyped || fromColumns?.trim()) {
      codes.push(column.code);
    }
  }
  return codes;
}

function laterDateWins(
  current: string | null | undefined,
  incoming: string | null | undefined,
): string | null {
  const a = current?.trim() || null;
  const b = incoming?.trim() || null;
  if (!a) return b;
  if (!b) return a;
  return new Date(b).getTime() >= new Date(a).getTime() ? b : a;
}

function eusrCategoryRowsFromRegisters(
  records: CustomerEusrRecord[],
): CustomerEusrCategoryCell[] {
  const byCategory = new Map<string, CustomerEusrCategoryCell>();
  for (const row of records) {
    const names = parseEusrCategories(row.eusrCategory);
    const categories = names.length > 0 ? names : ["EUSR"];
    for (const category of categories) {
      const existing = byCategory.get(category);
      const expiry = laterDateWins(existing?.expiry, row.expiry);
      const takeIncoming =
        !existing ||
        (row.expiry?.trim() && expiry === row.expiry?.trim());
      byCategory.set(category, {
        category,
        trainingDate: takeIncoming
          ? row.trainingDate
          : existing?.trainingDate ?? row.trainingDate,
        expiry,
      });
    }
  }
  return Array.from(byCategory.values());
}

function mergeEusrCategoryRows(
  matrix: MatrixSourceRow | null,
  registerRows: CustomerEusrCategoryCell[],
): CustomerEusrCategoryCell[] {
  const fromMatrix = listEusrCategoryMatrixCells(
    matrix?.columnValues,
    matrix?.categoryTrainingDates,
  ).map(({ category, trainingDate, expiry }) => ({
    category,
    trainingDate,
    expiry,
  }));
  if (fromMatrix.length === 0) return registerRows;
  const byCategory = new Map(
    fromMatrix.map((row) => [row.category.toLowerCase(), row] as const),
  );
  for (const row of registerRows) {
    const key = row.category.toLowerCase();
    const existing = byCategory.get(key);
    if (!existing) {
      byCategory.set(key, row);
      continue;
    }
    byCategory.set(key, {
      category: existing.category,
      trainingDate: existing.trainingDate ?? row.trainingDate,
      expiry: existing.expiry ?? row.expiry,
    });
  }
  return Array.from(byCategory.values());
}

function buildEnrichedRow(input: {
  candidate: WorkforceCandidate | null;
  matrix: MatrixSourceRow | null;
  nporsCategories: string[];
  nporsExpiry: string | null;
  nporsNumber?: string | null;
  swqrExpiry: string | null;
  eusrExpiry: string | null;
  eusrCategoryRows?: CustomerEusrCategoryCell[];
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
  const eusrCategoryRows = mergeEusrCategoryRows(
    matrix,
    input.eusrCategoryRows ?? [],
  );
  const nextExpiryDate =
    earliestExpiryDate([
      matrix?.nextExpiryDate,
      input.nporsExpiry,
      cscsExpiry,
      input.swqrExpiry ?? swqrFromMatrix,
      input.eusrExpiry ?? eusrFromMatrix,
      ...eusrCategoryRows.map((row) => row.expiry),
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
    eusrCategoryRows,
    inHouseCourse: input.inHouseCourse ?? null,
    inHouseExpiry: input.inHouseExpiry,
    ...dates,
    columnValues: matrix?.columnValues,
    categoryTrainingDates: matrix?.categoryTrainingDates,
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

  // With customer context: join company workforce to Training Matrix Update by
  // candidate name. Do NOT use listAdminMatrix(company) — that depends on a full
  // workforce reload and was dropping matrix dates (workforce-only UI).
  if (context) {
    const [workforce, exampleRows, npors, eusr, streetworks, inHouse] =
      await Promise.all([
        getAllowedWorkforceForCustomer(context),
        listTrainingMatrixExampleRows(),
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

    // Join by WorkforceItemId scoped to THIS company's allowed candidates. A
    // name-only join over the whole list would show one company's training on a
    // same-named candidate at another company.
    const allowedIds = new Set(
      workforce.map((row) => String(row.id).trim()).filter(Boolean),
    );
    const matrixById = new Map<string, MatrixSourceRow>();
    // Legacy rows written before the link columns existed have no owner id.
    // They are matched by name, but only when the name is unique across the
    // ENTIRE list — a duplicate name disqualifies every row with that name.
    const legacyByName = new Map<string, MatrixSourceRow | null>();
    for (const example of exampleRows) {
      const status = example.matrixLinkStatus?.trim();
      // Unresolved uploads never reach the customer portal (Task C).
      if (status === "Orphan" || status === "Needs Review") continue;

      const ownerId = String(example.workforceItemId ?? "").trim();
      if (ownerId) {
        if (!allowedIds.has(ownerId) || matrixById.has(ownerId)) continue;
        matrixById.set(
          ownerId,
          exampleToMatrixSource(example, context.companyName),
        );
        continue;
      }

      const key = nameKey(example.candidateName);
      if (!key) continue;
      // A legacy row that already carries a company link must belong to
      // THIS company, or it's excluded outright — never let a known
      // other-company row leak in (or falsely poison the name-uniqueness
      // check) just because a same-named candidate exists here too.
      const legacyCompanyId = String(example.companyItemId ?? "").trim();
      if (legacyCompanyId && legacyCompanyId !== context.companyId) continue;
      legacyByName.set(
        key,
        legacyByName.has(key)
          ? null
          : exampleToMatrixSource(example, context.companyName),
      );
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
    const eusrRowsByName = new Map<string, CustomerEusrRecord[]>();
    for (const row of eusr) {
      const key = nameKey(row.candidateName);
      if (!key) continue;
      const grouped = eusrRowsByName.get(key) ?? [];
      grouped.push(row);
      eusrRowsByName.set(key, grouped);
      if (!row.expiry?.trim()) continue;
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

    const rows: CustomerMatrixRecord[] = [];

    const workforceNameCounts = new Map<string, number>();
    for (const candidate of workforce) {
      const key = nameKey(candidate.candidateName);
      if (key) {
        workforceNameCounts.set(key, (workforceNameCounts.get(key) ?? 0) + 1);
      }
    }

    for (const candidate of workforce) {
      const key = nameKey(candidate.candidateName);
      const matrix =
        matrixById.get(String(candidate.id).trim()) ??
        ((workforceNameCounts.get(key) ?? 0) === 1
          ? (legacyByName.get(key) ?? null)
          : null);

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
          eusrCategoryRows: eusrCategoryRowsFromRegisters(
            eusrRowsByName.get(key) ?? [],
          ),
          inHouseExpiry: earliestExpiryDate([
            ...(inHouseExpiryByName.get(key) ?? []),
            matrix?.n031Expiry,
            matrix?.columnValues?.["N031 - Asbestos Awareness"] ?? null,
          ]),
          inHouseCourse: inHouseCourseByName.get(key) ?? null,
        }),
      );
    }

    const allowedNames = await getAllowedCandidateNames(context);
    return filterRowsByCandidateAccess(rows, allowedNames, context).sort(
      (a, b) => a.candidateName.localeCompare(b.candidateName),
    );
  }

  // Legacy / no context: matrix-only rows for the company name filter.
  const matrixRows = await listAdminMatrix(companyName);
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
      eusrCategoryRows: [],
      inHouseExpiry: null,
    }),
  );
}

/**
 * Dashboard stats — only loads lists needed for the cards.
 * Training-register counts are deferred (open Training Records for detail)
 * so first paint does not wait on NPORS/EUSR/Streetworks/In-House.
 */
export interface CustomerDashboardContent {
  stats: DashboardStats;
  offers: CustomerOfferRecord[];
  upcomingEvents: CustomerEventRecord[];
  /** Most recently uploaded documents, newest first — a slice of what's already fetched for documentsCount. */
  recentDocuments: CustomerDocumentRecord[];
  /** Candidates land on their own profile; everyone else lands on the candidate list. */
  profileShortcut: { href: string; label: string };
}

export async function getCustomerDashboardContent(
  context: CustomerContext,
): Promise<CustomerDashboardContent> {
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

  const stats: DashboardStats = {
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

  const upcomingEvents = events.filter((event) => {
    if (!event.eventDate) return false;
    const time = new Date(event.eventDate).getTime();
    return !Number.isNaN(time) && time >= now;
  });

  const recentDocuments = [...documents]
    .sort((a, b) => {
      const timeA = a.uploadedDate ? new Date(a.uploadedDate).getTime() : NaN;
      const timeB = b.uploadedDate ? new Date(b.uploadedDate).getTime() : NaN;
      if (Number.isNaN(timeA) && Number.isNaN(timeB)) return 0;
      if (Number.isNaN(timeA)) return 1;
      if (Number.isNaN(timeB)) return -1;
      return timeB - timeA;
    })
    .slice(0, 5);

  const profileShortcut =
    context.customerRole === "Candidate" && workforce.length === 1
      ? { href: `/customer/candidates/${workforce[0].id}`, label: "View my profile" }
      : { href: "/customer/candidates", label: "View candidates" };

  return { stats, offers, upcomingEvents, recentDocuments, profileShortcut };
}

/** Backwards-compatible stats-only dashboard service. */
export async function getCustomerDashboard(
  context: CustomerContext,
): Promise<DashboardStats> {
  return (await getCustomerDashboardContent(context)).stats;
}

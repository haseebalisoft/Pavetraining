import { redirect } from "next/navigation";

import { CandidateProfileView } from "@/components/customer/CandidateProfileView";
import { auth } from "@/auth";
import {
  listAdminDocuments,
  listAdminMatrix,
  listAdminNvq,
  listAdminRegister,
  type AdminMatrixRecord,
  type AdminNvqRecord,
  type AdminTrainingRecord,
} from "@/lib/services/adminCrudService";
import type { AdminDocumentRecord } from "@/types/adminDocuments";
import { toCustomerOutcome } from "@/lib/training/customerOutcome";
import { earliestExpiryDate } from "@/lib/training/expiryFilters";
import { getWorkforceById } from "@/lib/services/workforceService";
import type {
  CustomerDocumentRecord,
  CustomerEusrRecord,
  CustomerInHouseRecord,
  CustomerMatrixRecord,
  CustomerNporsRecord,
  CustomerNvqRecord,
  CustomerStreetworksRecord,
} from "@/types/models";

export const dynamic = "force-dynamic";

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

function matchesCandidate(
  row: {
    candidateName?: string | null;
    candidate?: string | null;
    candidateLookupId?: string | null;
  },
  candidateName: string,
  candidateId?: string,
): boolean {
  if (candidateId && row.candidateLookupId === candidateId) return true;
  const key = nameKey(candidateName);
  if (!key) return false;
  return nameKey(row.candidateName) === key || nameKey(row.candidate) === key;
}

function mapMatrix(
  row: AdminMatrixRecord | null,
  candidateId: string,
): CustomerMatrixRecord | null {
  if (!row) return null;
  const nporsExpiry = earliestExpiryDate([
    row.n001Expiry,
    row.n003Expiry,
    row.n004Expiry,
    row.n010Expiry,
    row.n020Expiry,
    row.n021Expiry,
    row.n027Expiry,
    row.n100Expiry,
  ]);
  return {
    id: row.id,
    candidateId,
    candidateName: row.candidateName,
    companyName: row.companyName,
    dateOfBirth: row.dateOfBirth,
    department: row.department,
    trainingManager: null,
    supervisor: null,
    overallStatus: row.overallStatus,
    needsReview: row.needsReview,
    nextExpiryDate: row.nextExpiryDate,
    nporsCategories: null,
    nporsExpiry,
    nporsNumber: null,
    cscsNumber: null,
    cscsExpiry: row.cscsExpiry ?? row.columnValues?.["CSCS Expiry"] ?? null,
    swqrNumber: null,
    swqrExpiry: row.columnValues?.["NRSWA Expiry"] ?? null,
    eusrNumber: null,
    eusrExpiry: row.columnValues?.["EUSR Expiry"] ?? null,
    inHouseCourse: null,
    inHouseExpiry: row.n031Expiry ?? row.columnValues?.["N031 - Asbestos Awareness"] ?? null,
    n001Expiry: row.n001Expiry,
    n003Expiry: row.n003Expiry,
    n004Expiry: row.n004Expiry,
    n010Expiry: row.n010Expiry,
    n020Expiry: row.n020Expiry,
    n021Expiry: row.n021Expiry,
    n027Expiry: row.n027Expiry,
    n100Expiry: row.n100Expiry,
  };
}

function mapNpors(
  row: AdminTrainingRecord,
  workforceId: string,
): CustomerNporsRecord {
  return {
    id: row.id,
    candidateName: row.candidateName,
    workforceId,
    nporsNumber: row.nporsNumber ?? null,
    trainingDate: row.trainingDate,
    trainingAddress: row.trainingAddress,
    noviceOrEwt: row.noviceOrEwt ?? null,
    nporsCategory: row.nporsCategory ?? null,
    outcome: toCustomerOutcome(row.trainingOutcome),
    expiry: row.expiry,
  };
}

function mapEusr(
  row: AdminTrainingRecord,
  workforceId: string,
): CustomerEusrRecord {
  return {
    id: row.id,
    candidateName: row.candidateName,
    workforceId,
    eusrNumber: row.eusrNumber ?? null,
    eusrCategory: row.eusrCategory ?? null,
    cardType: row.cardType ?? null,
    trainingDate: row.trainingDate,
    trainingAddress: row.trainingAddress,
    outcome: toCustomerOutcome(row.trainingOutcome),
    expiry: row.expiry,
  };
}

function mapStreetworks(
  row: AdminTrainingRecord,
  workforceId: string,
): CustomerStreetworksRecord {
  return {
    id: row.id,
    candidateName: row.candidateName,
    workforceId,
    swqrNumber: row.swqrNumber ?? null,
    course: row.course ?? null,
    streetworksCategory: row.streetworksCategory ?? null,
    trainingDate: row.trainingDate,
    trainingDateEnd: row.trainingDateEnd ?? null,
    trainingAddress: row.trainingAddress,
    outcome: toCustomerOutcome(row.trainingOutcome),
    expiry: row.expiry,
  };
}

function mapInHouse(
  row: AdminTrainingRecord,
  workforceId: string,
): CustomerInHouseRecord {
  return {
    id: row.id,
    candidateName: row.candidateName,
    workforceId,
    course: row.certificateCategory ?? row.courseCategory ?? null,
    trainingDate: row.trainingDate,
    trainingAddress: row.trainingAddress,
    outcome: toCustomerOutcome(row.trainingOutcome),
    expiry: row.expiry,
  };
}

function mapNvq(row: AdminNvqRecord): CustomerNvqRecord {
  return {
    id: row.id,
    candidateName: row.candidateName,
    nvqTitle: row.nvqTitle,
    boltOn: row.boltOn,
    dateRegistered: row.dateRegistered,
    inductionDate: row.inductionDate,
    stageOfNvq: row.stageOfNvq,
    notes: row.notes,
    completedDate: row.completedDate,
    status: row.status,
    ulnNumber: row.ulnNumber,
    cardSchemeCategory: row.cardSchemeCategory,
    cardExtensionDateNeeded: row.cardExtensionDateNeeded,
  };
}

function mapDocument(row: AdminDocumentRecord): CustomerDocumentRecord {
  return {
    id: row.id,
    name: row.name,
    documentType: row.documentType,
    candidate: row.candidate,
    uploadedDate: row.modifiedDate ?? row.uploadedDate,
    canDownload: row.canDownload,
    viewPath: row.previewPath,
    downloadPath: row.downloadPath ?? `/api/admin/documents/${row.id}/download`,
  };
}

export default async function AdminCandidateProfilePage({
  params,
}: {
  params: Promise<{ candidateId: string }>;
}) {
  const session = await auth();
  if (!session?.user?.email) {
    redirect("/login");
  }

  const { candidateId } = await params;
  const candidate = await getWorkforceById(candidateId);
  if (!candidate) {
    redirect("/admin/workforce");
  }

  const [
    matrixRows,
    npors,
    eusr,
    streetworks,
    inHouse,
    nvq,
    documents,
  ] = await Promise.all([
    listAdminMatrix(candidate.companyName),
    listAdminRegister("nporsRegister", candidate.companyName),
    listAdminRegister("eusrRegister", candidate.companyName),
    listAdminRegister("nrswaRegister", candidate.companyName),
    listAdminRegister("inHouseCertificates", candidate.companyName),
    listAdminNvq(candidate.companyName),
    listAdminDocuments({ companyName: candidate.companyName }),
  ]);

  const matrixRow =
    matrixRows.find(
      (row) =>
        row.workforceId === candidateId ||
        nameKey(row.candidateName) === nameKey(candidate.candidateName),
    ) ?? null;

  const filterRows = <
    T extends {
      candidateName?: string | null;
      candidate?: string | null;
      candidateLookupId?: string | null;
    },
  >(
    rows: T[],
  ) =>
    rows.filter((row) =>
      matchesCandidate(row, candidate.candidateName, candidateId),
    );

  return (
    <CandidateProfileView
      variant="admin"
      candidate={candidate}
      matrixRow={mapMatrix(matrixRow, candidateId)}
      matrixReturnHref="/admin/workforce"
      nporsRecords={filterRows(npors).map((row) => mapNpors(row, candidateId))}
      eusrRecords={filterRows(eusr).map((row) => mapEusr(row, candidateId))}
      streetworksRecords={filterRows(streetworks).map((row) =>
        mapStreetworks(row, candidateId),
      )}
      inHouseRecords={filterRows(inHouse).map((row) =>
        mapInHouse(row, candidateId),
      )}
      nvqRecords={filterRows(nvq).map(mapNvq)}
      documents={filterRows(documents).map(mapDocument)}
    />
  );
}

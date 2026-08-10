import "server-only";

import type { AdminRegisterKey, AdminTrainingRecord } from "@/lib/services/adminCrudService";
import type { MatrixSyncResult, MatrixSyncResultItem } from "@/types/matrixSync";
import {
  syncAfterRegisterDelete,
  syncAfterRegisterSave,
} from "@/lib/services/trainingMatrixSyncService";

/**
 * Best-effort candidate-only matrix sync after a register write.
 * Never fails the primary create/update response.
 */
export async function triggerMatrixSyncAfterRegister(
  registerKey: AdminRegisterKey,
  record: AdminTrainingRecord,
  userEmail?: string | null,
): Promise<MatrixSyncResult | null> {
  try {
    return await syncAfterRegisterSave(registerKey, record, {
      dryRun: false,
      userEmail,
    });
  } catch (error) {
    console.error(
      `[matrix-sync] Failed after ${registerKey} save #${record.id}`,
      error,
    );
    try {
      const { sendAdminAlert } = await import(
        "@/lib/services/notificationService"
      );
      await sendAdminAlert({
        title: "Training Matrix sync error",
        detail:
          error instanceof Error
            ? `${registerKey} #${record.id}: ${error.message}`
            : `Matrix sync failed after ${registerKey} #${record.id}`,
        itemId: record.id,
        actorEmail: userEmail,
      });
    } catch {
      // ignore alert failures
    }
    return null;
  }
}

/**
 * Best-effort candidate-only matrix recompute after a register record is
 * deleted. Never fails the primary delete response.
 */
export async function triggerMatrixSyncAfterRegisterDelete(
  registerKey: AdminRegisterKey,
  deletedRecord: AdminTrainingRecord,
  userEmail?: string | null,
): Promise<MatrixSyncResult | null> {
  try {
    return await syncAfterRegisterDelete(registerKey, deletedRecord, {
      dryRun: false,
      userEmail,
    });
  } catch (error) {
    console.error(
      `[matrix-sync] Failed to recompute after ${registerKey} delete #${deletedRecord.id}`,
      error,
    );
    try {
      const { sendAdminAlert } = await import(
        "@/lib/services/notificationService"
      );
      await sendAdminAlert({
        title: "Training Matrix delete-recompute error",
        detail:
          error instanceof Error
            ? `${registerKey} #${deletedRecord.id}: ${error.message}`
            : `Matrix recompute failed after ${registerKey} #${deletedRecord.id} delete`,
        itemId: deletedRecord.id,
        actorEmail: userEmail,
      });
    } catch {
      // ignore alert failures
    }
    return null;
  }
}

/**
 * NVQ has no Training Matrix expiry/category target — its fields (card
 * scheme, ULN, stage, card extension date, …) are unrelated to the NPORS/
 * EUSR/NRSWA/N031 Asbestos columns. Per product decision, NVQ save/update
 * never attempts a Matrix write, so it can never fail one; this always
 * returns the same fixed, non-error note through the identical MatrixSyncResult
 * shape the other registers use, so AdminCrudPage's existing toast path
 * surfaces it with no NVQ-specific UI code.
 */
export function triggerMatrixSyncAfterNvq(record: {
  candidateName: string;
  companyName?: string | null;
}): MatrixSyncResult {
  const items: MatrixSyncResultItem[] = [
    {
      candidate: record.candidateName,
      company: record.companyName ?? "—",
      candidateId: null,
      companyId: null,
      registerSources: [],
      matrixRowId: null,
      matrixRowFound: false,
      matrixRowCreated: false,
      fieldsUpdated: [],
      warnings: [],
      errors: [],
      skipped: true,
      skipReason:
        "NVQ saved to profile/history; no Matrix expiry target configured.",
    },
  ];
  return {
    dryRun: false,
    scope: "register",
    items,
    summary: {
      processed: 1,
      updated: 0,
      created: 0,
      skipped: 1,
      errors: 0,
      warnings: 0,
    },
  };
}

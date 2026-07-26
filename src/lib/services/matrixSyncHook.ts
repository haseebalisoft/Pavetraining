import "server-only";

import type { AdminRegisterKey, AdminTrainingRecord } from "@/lib/services/adminCrudService";
import type { MatrixSyncResult } from "@/types/matrixSync";
import {
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

import { WebPartContext } from "@microsoft/sp-webpart-base";

import type { RegisterRecord, TrainingTrackerStats } from "../models";
import { NporsService } from "./NporsService";

function parseDate(value: string | null | undefined): Date | null {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function isActiveStatus(status: string): boolean {
  const normalized = (status || "").trim().toLowerCase();
  if (!normalized) return true;
  if (
    normalized.indexOf("fail") >= 0 ||
    normalized.indexOf("inactive") >= 0 ||
    normalized.indexOf("withdraw") >= 0 ||
    normalized.indexOf("cancel") >= 0
  ) {
    return false;
  }
  return true;
}

function isCompletedStatus(status: string): boolean {
  const normalized = (status || "").trim().toLowerCase();
  return (
    normalized.indexOf("pass") >= 0 ||
    normalized.indexOf("complete") >= 0 ||
    normalized.indexOf("achieved") >= 0
  );
}

function startOfMonth(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function endOfMonth(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59, 999);
}

function countExpiringInDays(rows: RegisterRecord[], days: number): number {
  const now = new Date();
  const horizon = new Date();
  horizon.setDate(horizon.getDate() + days);

  let count = 0;
  for (let i = 0; i < rows.length; i++) {
    const expiry = parseDate(rows[i].expiryDate);
    if (!expiry) continue;
    if (expiry >= now && expiry <= horizon) {
      count += 1;
    }
  }
  return count;
}

function countActive(rows: RegisterRecord[]): number {
  let count = 0;
  for (let i = 0; i < rows.length; i++) {
    if (isActiveStatus(rows[i].status)) {
      count += 1;
    }
  }
  return count;
}

function countCompletedThisMonth(rows: RegisterRecord[]): number {
  const now = new Date();
  const from = startOfMonth(now);
  const to = endOfMonth(now);
  let count = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!isCompletedStatus(row.status)) continue;
    const when = parseDate(row.expiryDate);
    if (!when) continue;
    if (when >= from && when <= to) {
      count += 1;
    }
  }
  return count;
}

/**
 * Training Tracker band metrics derived from register + workforce lists.
 * Never throws — returns zeros on failure.
 */
export class StatsService {
  public static async getTrainingTrackerStats(
    context: WebPartContext
  ): Promise<TrainingTrackerStats> {
    try {
      const [workforce, registers] = await Promise.all([
        NporsService.getTopWorkforce(context, 500),
        NporsService.getRegisterRowsForStats(context, 500),
      ]);

      const allRegisters: RegisterRecord[] = ([] as RegisterRecord[])
        .concat(registers.npors)
        .concat(registers.eusr)
        .concat(registers.nrswa)
        .concat(registers.inHouse)
        .concat(registers.nvq);

      const activeWorkforce = workforce.filter((member) => {
        return Boolean(member.name);
      });

      return {
        totalOperators:
          activeWorkforce.length > 0
            ? activeWorkforce.length
            : new Set(
                allRegisters
                  .map((row) => row.operator.trim().toLowerCase())
                  .filter(Boolean)
              ).size,
        expiringIn30Days: countExpiringInDays(allRegisters, 30),
        activeRegistrations: countActive(allRegisters),
        completedThisMonth: countCompletedThisMonth(allRegisters),
      };
    } catch (error) {
      console.error(
        "[StatsService.getTrainingTrackerStats] Failed to compute stats",
        error
      );
      throw error;
    }
  }
}

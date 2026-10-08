"use client";

import { useEffect, useId, useState } from "react";

import { useAdminToast } from "@/components/admin/AdminToast";
import styles from "@/components/admin/admin.module.css";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { AdminWorkforceRecord } from "@/lib/services/adminCrudService";

export type WorkforceInlineField =
  | "department"
  | "trainingManager"
  | "supervisor";

export type WorkforceInlineOption = {
  value: string;
  label: string;
  /** Optional email used when saving Training Manager / Supervisor. */
  email?: string;
};

/**
 * Compact company-scoped select for Workforce table cells. Saves on change via
 * PATCH /api/admin/workforce/[id] and rolls back if SharePoint rejects it.
 */
export function WorkforceInlineSelect({
  row,
  field,
  options,
  onSaved,
}: {
  row: AdminWorkforceRecord;
  field: WorkforceInlineField;
  options: WorkforceInlineOption[];
  onSaved: (record: AdminWorkforceRecord) => void;
}) {
  const labelId = useId();
  const { pushToast } = useAdminToast();
  const currentValue = String(row[field] ?? "").trim();
  const [value, setValue] = useState(currentValue);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    setValue(currentValue);
  }, [currentValue, row.id]);

  const hasCompany = Boolean(
    String(row.companyId ?? "").trim() || String(row.companyName ?? "").trim(),
  );

  // Ensure the current assignment stays selectable even if it fell out of the
  // Active / company filter (same behaviour as the edit drawer).
  const selectOptions = (() => {
    if (!hasCompany) {
      return [{ value: "", label: "No company on this row" }];
    }
    const key = currentValue.toLowerCase();
    const hasCurrent =
      !currentValue ||
      options.some((option) => option.value.trim().toLowerCase() === key);
    const merged =
      currentValue && !hasCurrent
        ? [{ value: currentValue, label: `${currentValue} (current)` }, ...options]
        : options;
    if (merged.length === 0) {
      return [
        {
          value: "",
          label:
            field === "department"
              ? "No departments for this company"
              : field === "trainingManager"
                ? "No Training Managers for this company"
                : "No Supervisors for this company",
        },
      ];
    }
    return [{ value: "", label: "— None —" }, ...merged];
  })();

  async function commit(nextRaw: string) {
    const next = nextRaw.trim();
    if (next === value.trim()) return;
    if (!hasCompany) {
      pushToast("This candidate has no company — set company in Edit first.", "error");
      return;
    }

    const previous = value;
    setValue(next);
    setSaving(true);
    try {
      const picked = options.find(
        (option) => option.value.trim().toLowerCase() === next.toLowerCase(),
      );
      const body: Record<string, unknown> = {
        [field]: next,
        companyId: row.companyId ?? "",
        companyName: row.companyName ?? "",
      };
      if (field === "trainingManager" && picked?.email) {
        body.trainingManagerEmail = picked.email;
      }
      if (field === "supervisor" && picked?.email) {
        body.supervisorEmail = picked.email;
      }

      const response = await fetch(`/api/admin/workforce/${row.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      const payload = (await response.json().catch(() => null)) as {
        record?: AdminWorkforceRecord;
      } | null;
      if (!payload?.record) {
        throw new Error("Server did not return the updated candidate.");
      }
      onSaved(payload.record);
    } catch (error) {
      setValue(previous);
      pushToast(
        error instanceof Error ? error.message : "Could not save change.",
        "error",
      );
    } finally {
      setSaving(false);
    }
  }

  const fieldLabel =
    field === "department"
      ? "Department"
      : field === "trainingManager"
        ? "Training manager"
        : "Supervisor";

  return (
    <div className={styles.inlineSelectWrap}>
      <label className={styles.srOnly} htmlFor={labelId}>
        {fieldLabel} for {row.candidateName || row.id}
      </label>
      <select
        id={labelId}
        className={styles.inlineTableSelect}
        value={value}
        disabled={saving || !hasCompany}
        aria-busy={saving || undefined}
        onChange={(event) => {
          void commit(event.target.value);
        }}
      >
        {selectOptions.map((option, index) => (
          <option
            key={`${field}-${option.value}-${index}`}
            value={option.value}
          >
            {option.label}
          </option>
        ))}
      </select>
      {saving ? (
        <span className={styles.inlineSelectSaving} aria-live="polite">
          Saving…
        </span>
      ) : null}
    </div>
  );
}

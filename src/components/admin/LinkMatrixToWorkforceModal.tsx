"use client";

import { useEffect, useMemo, useState } from "react";

import { readPublicApiError } from "@/lib/errors/publicMessages";
import type {
  AdminMatrixRecord,
  AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";

import styles from "./admin.module.css";

/**
 * Admin repair modal shown when an Orphan / Needs Review matrix row needs a
 * Workforce owner manually assigned. Loads Workforce candidates (scoped to
 * the row's company when available), lets the admin pick one, and calls
 * POST /api/admin/training-matrix/[id]/link.
 */
export interface LinkMatrixToWorkforceModalProps {
  open: boolean;
  matrixRow: AdminMatrixRecord | null;
  onClose: () => void;
  onLinked: (record: AdminMatrixRecord) => void;
}

type WorkforceOption = Pick<
  AdminWorkforceRecord,
  "id" | "candidateName" | "companyName" | "workforceNumber" | "dateOfBirth"
>;

export function LinkMatrixToWorkforceModal({
  open,
  matrixRow,
  onClose,
  onLinked,
}: LinkMatrixToWorkforceModalProps) {
  const [candidates, setCandidates] = useState<WorkforceOption[]>([]);
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState("");
  const [selectedId, setSelectedId] = useState<string>("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    setSearch("");
    setSelectedId("");
    setError(null);
    let cancelled = false;
    setLoading(true);
    fetch("/api/admin/workforce", { cache: "no-store" })
      .then(async (response) => {
        if (!response.ok) throw new Error(await readPublicApiError(response));
        const payload = (await response.json().catch(() => null)) as {
          records?: WorkforceOption[];
        } | null;
        return payload?.records ?? [];
      })
      .then((records) => {
        if (cancelled) return;
        setCandidates(records);
      })
      .catch((err) => {
        if (cancelled) return;
        setError(
          err instanceof Error ? err.message : "Failed to load Workforce.",
        );
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [open]);

  useEffect(() => {
    if (!open) return;
    function onKey(event: KeyboardEvent) {
      if (event.key === "Escape" && !saving) onClose();
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, saving, onClose]);

  const targetCompanyKey =
    matrixRow?.companyName?.trim().toLowerCase() ?? "";

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    // Prefer candidates from the same company; fall back to all when the
    // matrix row has no company yet (Orphan without company link).
    const scoped = targetCompanyKey
      ? candidates.filter(
          (row) =>
            (row.companyName ?? "").trim().toLowerCase() === targetCompanyKey,
        )
      : candidates;
    if (!query) return scoped.slice(0, 200);
    return scoped
      .filter((row) => {
        const name = (row.candidateName ?? "").toLowerCase();
        const number = (row.workforceNumber ?? "").toLowerCase();
        return name.includes(query) || number.includes(query);
      })
      .slice(0, 200);
  }, [candidates, targetCompanyKey, search]);

  if (!open || !matrixRow) return null;

  async function onSubmit() {
    if (!matrixRow) return;
    if (!selectedId) {
      setError("Pick a Workforce candidate to link to.");
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const response = await fetch(
        `/api/admin/training-matrix/${encodeURIComponent(matrixRow.id)}/link`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ workforceId: selectedId }),
        },
      );
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const payload = (await response.json().catch(() => null)) as {
        record?: AdminMatrixRecord;
      } | null;
      if (!payload?.record) throw new Error("Server did not return the linked row.");
      onLinked(payload.record);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to link.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div
      className={styles.quickAddScrim}
      role="dialog"
      aria-modal="true"
      aria-label="Link matrix row to Workforce candidate"
      onClick={(event) => {
        if (event.target === event.currentTarget && !saving) onClose();
      }}
    >
      <div className={styles.quickAddPanel}>
        <header className={styles.quickAddHeader}>
          <h3 className={styles.quickAddTitle}>
            Link to Workforce · {matrixRow.candidateName}
          </h3>
          <button
            type="button"
            className={styles.quickAddClose}
            onClick={onClose}
            disabled={saving}
            aria-label="Close"
          >
            ×
          </button>
        </header>
        <div className={styles.quickAddForm}>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Search Workforce</span>
            <input
              className={styles.input}
              type="search"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Name or Workforce Number"
              disabled={saving}
              autoFocus
            />
          </label>
          <p className={styles.quickAddHelp}>
            {targetCompanyKey
              ? `Showing candidates in ${matrixRow.companyName}. Pick one to link this matrix row to.`
              : "This matrix row has no company yet — showing all Workforce candidates."}
          </p>
          <div className={styles.linkMatrixList}>
            {loading ? (
              <p className={styles.quickAddHelp}>Loading Workforce…</p>
            ) : filtered.length === 0 ? (
              <p className={styles.quickAddHelp}>
                No Workforce candidates match. Add the candidate under Admin
                → Workforce first, then re-open this dialog.
              </p>
            ) : (
              filtered.map((row) => (
                <label key={row.id} className={styles.linkMatrixOption}>
                  <input
                    type="radio"
                    name="linkWorkforce"
                    value={row.id}
                    checked={selectedId === row.id}
                    onChange={() => setSelectedId(row.id)}
                    disabled={saving}
                  />
                  <span className={styles.linkMatrixOptionText}>
                    <strong>{row.candidateName}</strong>
                    <span>
                      {row.companyName ?? "—"} · #{row.workforceNumber ?? "—"}
                      {row.dateOfBirth ? ` · DOB ${row.dateOfBirth}` : ""}
                    </span>
                  </span>
                </label>
              ))
            )}
          </div>
          {error ? (
            <p className={styles.quickAddError} role="alert">
              {error}
            </p>
          ) : null}
          <div className={styles.quickAddActions}>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={onClose}
              disabled={saving}
            >
              Cancel
            </button>
            <button
              type="button"
              className={styles.primaryButton}
              disabled={saving || !selectedId}
              onClick={() => void onSubmit()}
            >
              {saving ? "Linking…" : "Link to Workforce"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

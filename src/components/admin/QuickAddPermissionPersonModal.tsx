"use client";

import { useEffect, useRef, useState, type FormEvent } from "react";

import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { AdminPermissionRecord } from "@/lib/services/adminCrudService";

import styles from "./admin.module.css";

/**
 * Minimal dialog for adding a Training Manager / Supervisor without leaving
 * the current form (Workforce create/edit). Company is passed in and read-only
 * so the new Permissions row is always scoped to the current context. Posts
 * to /api/admin/permissions and hands the saved record back to the caller so
 * it can refresh its dropdown + auto-select the new person.
 */
export interface QuickAddPermissionPersonModalProps {
  open: boolean;
  onClose: () => void;
  onCreated: (record: AdminPermissionRecord) => void;
  /**
   * Which SharePoint RoleType to create. The modal preselects the underlying
   * `permissionRole` form value: Training Manager → Manager, Supervisor →
   * Supervisor. Candidate is included for symmetry with the strict filter but
   * not used by Workforce today.
   */
  role: "Training Manager" | "Supervisor" | "Candidate";
  companyId: string;
  companyName: string;
}

function permissionRoleFor(
  role: QuickAddPermissionPersonModalProps["role"],
): "Admin" | "Manager" | "Supervisor" | "Customer" | "Candidate" {
  if (role === "Training Manager") return "Manager";
  if (role === "Supervisor") return "Supervisor";
  return "Candidate";
}

export function QuickAddPermissionPersonModal({
  open,
  onClose,
  onCreated,
  role,
  companyId,
  companyName,
}: QuickAddPermissionPersonModalProps) {
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const nameRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (open) {
      setName("");
      setEmail("");
      setError(null);
      setSaving(false);
      // Focus after the modal has actually rendered.
      const timer = window.setTimeout(() => nameRef.current?.focus(), 20);
      return () => window.clearTimeout(timer);
    }
    return undefined;
  }, [open]);

  useEffect(() => {
    if (!open) return;
    function onKey(event: KeyboardEvent) {
      if (event.key === "Escape") onClose();
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) return null;

  async function onSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const trimmedName = name.trim();
    const trimmedEmail = email.trim();
    if (!trimmedName) {
      setError("Name is required.");
      return;
    }
    if (!trimmedEmail) {
      setError("Email is required.");
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const response = await fetch("/api/admin/permissions", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: trimmedName,
          userEmail: trimmedEmail,
          permissionRole: permissionRoleFor(role),
          companyId,
          companyName,
          status: "Active",
          accessScope: "Full Company",
          canView: true,
          canDownload: false,
          canEdit: false,
        }),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      const payload = (await response.json().catch(() => null)) as {
        record?: AdminPermissionRecord;
      } | null;
      if (!payload?.record) {
        throw new Error("Server did not return the new Permissions row.");
      }
      onCreated(payload.record);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to add.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div
      className={styles.quickAddScrim}
      role="dialog"
      aria-modal="true"
      aria-label={`Add ${role}`}
      onClick={(event) => {
        if (event.target === event.currentTarget && !saving) onClose();
      }}
    >
      <div className={styles.quickAddPanel}>
        <header className={styles.quickAddHeader}>
          <h3 className={styles.quickAddTitle}>Add {role}</h3>
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
        <form onSubmit={onSubmit} className={styles.quickAddForm}>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Name</span>
            <input
              ref={nameRef}
              className={styles.input}
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              disabled={saving}
              autoComplete="off"
              required
            />
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Email</span>
            <input
              className={styles.input}
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              disabled={saving}
              autoComplete="off"
              required
            />
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Company</span>
            <input
              className={styles.input}
              type="text"
              value={companyName}
              readOnly
              disabled
            />
          </label>
          <p className={styles.quickAddHelp}>
            Role: <strong>{role}</strong> · Status: <strong>Active</strong> ·
            Access scope: <strong>Full Company</strong>
          </p>
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
              type="submit"
              className={styles.primaryButton}
              disabled={saving}
            >
              {saving ? "Saving…" : `Save ${role}`}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

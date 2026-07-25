"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { motion } from "framer-motion";

import { AdminDrawer } from "@/components/admin/AdminDrawer";
import {
  DocumentsBrowseView,
  type BrowsePath,
  companyKeyFromRecord,
  typeKeyFromRecord,
} from "@/components/admin/documents/DocumentsBrowseView";
import styles from "@/components/admin/admin.module.css";
import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { AdminDocumentRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

type FormState = Record<string, string | boolean>;

const fields = [
  { name: "name", label: "File name", type: "text" as const, required: true },
  { name: "company", label: "Company", type: "company" as const },
  { name: "candidate", label: "Candidate", type: "text" as const },
  { name: "documentType", label: "Document type", type: "text" as const },
  {
    name: "customerVisible",
    label: "Customer visible",
    type: "boolean" as const,
  },
  {
    name: "notificationSent",
    label: "Notification sent",
    type: "boolean" as const,
  },
];

function buildForm(row: AdminDocumentRecord): FormState {
  return {
    name: row.name ?? "",
    company: row.company ?? "",
    candidate: row.candidate ?? "",
    documentType: row.documentType ?? "",
    customerVisible: Boolean(row.customerVisible),
    notificationSent: Boolean(row.notificationSent),
  };
}

function validate(form: FormState): string | null {
  for (const field of fields) {
    if (!("required" in field && field.required)) continue;
    const value = form[field.name];
    if (value === "" || value === null || value === undefined) {
      return `${field.label} is required.`;
    }
  }
  return null;
}

/**
 * FLAG: AdminDocumentsClient keeps list/reload/edit-drawer logic here because
 * AdminCrudPage couples table markup with that lifecycle. Folder browse UI is
 * extracted to DocumentsBrowseView + sibling presentational components.
 */
export function AdminDocumentsClient({
  companies,
  initialRows,
}: {
  companies: Company[];
  initialRows: AdminDocumentRecord[];
}) {
  const { pushToast } = useAdminToast();
  const [rows, setRows] = useState(initialRows);
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState("");
  const [documentTypeFilter, setDocumentTypeFilter] = useState("");
  const [candidateFilter, setCandidateFilter] = useState("");
  const [visibilityFilter, setVisibilityFilter] = useState("");
  const [path, setPath] = useState<BrowsePath>({ level: "companies" });
  const [direction, setDirection] = useState(1);
  const [busyId, setBusyId] = useState<string | null>(null);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editing, setEditing] = useState<AdminDocumentRecord | null>(null);
  const [form, setForm] = useState<FormState>({});
  const [formError, setFormError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  const documentTypes = useMemo(() => {
    const values = new Set<string>();
    for (const row of rows) {
      if (row.documentType?.trim()) values.add(row.documentType.trim());
    }
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const candidates = useMemo(() => {
    const values = new Set<string>();
    for (const row of rows) {
      if (row.candidate?.trim()) values.add(row.candidate.trim());
    }
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const response = await fetch("/api/admin/documents", { cache: "no-store" });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      const payload = (await response.json()) as {
        records?: AdminDocumentRecord[];
      };
      setRows(payload.records ?? []);
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to load documents",
        "error",
      );
    } finally {
      setLoading(false);
    }
  }, [pushToast]);

  const filteredRows = useMemo(() => {
    const query = search.trim().toLowerCase();
    return rows.filter((row) => {
      if (
        documentTypeFilter &&
        (row.documentType ?? "").trim().toLowerCase() !==
          documentTypeFilter.trim().toLowerCase()
      ) {
        return false;
      }
      if (
        candidateFilter &&
        (row.candidate ?? "").trim().toLowerCase() !==
          candidateFilter.trim().toLowerCase()
      ) {
        return false;
      }
      if (visibilityFilter === "yes" && !row.customerVisible) return false;
      if (visibilityFilter === "no" && row.customerVisible) return false;
      if (!query) return true;
      const haystack = [
        row.name,
        row.documentType,
        row.company,
        row.candidate,
        row.metadataStatus,
        row.modifiedBy,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return haystack.includes(query);
    });
  }, [
    rows,
    search,
    documentTypeFilter,
    candidateFilter,
    visibilityFilter,
  ]);

  // Keep breadcrumb path valid when filters shrink the visible set.
  useEffect(() => {
    if (path.level === "companies") return;

    const companyStillExists = filteredRows.some(
      (row) => companyKeyFromRecord(row) === path.companyKey,
    );
    if (!companyStillExists) {
      setDirection(-1);
      setPath({ level: "companies" });
      return;
    }

    if (path.level === "files") {
      const typeStillExists = filteredRows.some(
        (row) =>
          companyKeyFromRecord(row) === path.companyKey &&
          typeKeyFromRecord(row) === path.typeKey,
      );
      if (!typeStillExists) {
        setDirection(-1);
        setPath({
          level: "types",
          companyKey: path.companyKey,
          companyLabel: path.companyLabel,
        });
      }
    }
  }, [filteredRows, path]);

  function navigate(next: BrowsePath, nextDirection: -1 | 1) {
    setDirection(nextDirection);
    setPath(next);
  }

  function openEdit(row: AdminDocumentRecord) {
    setEditing(row);
    setForm(buildForm(row));
    setFormError(null);
    setDrawerOpen(true);
  }

  async function setVisibility(
    row: AdminDocumentRecord,
    customerVisible: boolean,
  ) {
    setBusyId(row.id);
    try {
      const response = await fetch(`/api/admin/documents/${row.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ customerVisible }),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      pushToast(
        customerVisible
          ? "Document marked customer visible."
          : "Document hidden from customer.",
        "success",
      );
      await load();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to update visibility",
        "error",
      );
    } finally {
      setBusyId(null);
    }
  }

  async function persist() {
    if (!editing) return;
    const error = validate(form);
    if (error) {
      setFormError(error);
      return;
    }

    setSaving(true);
    setFormError(null);
    try {
      const response = await fetch(`/api/admin/documents/${editing.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      pushToast("Record updated.");
      setDrawerOpen(false);
      await load();
    } catch (error) {
      const message = error instanceof Error ? error.message : "Save failed.";
      setFormError(message);
      pushToast(message, "error");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={[
              { label: "Admin", href: "/admin" },
              { label: "Documents" },
            ]}
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Documents</h1>
          <p className={styles.subtitle}>
            Browse by company and document type. Edit metadata, visibility, and
            company links without changing the underlying SharePoint library.
          </p>
        </div>
      </header>

      <div className={styles.crudToolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            value={search}
            placeholder="Search documents…"
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Document type</span>
          <select
            className={styles.select}
            value={documentTypeFilter}
            onChange={(event) => setDocumentTypeFilter(event.target.value)}
          >
            <option value="">All types</option>
            {documentTypes.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Candidate</span>
          <select
            className={styles.select}
            value={candidateFilter}
            onChange={(event) => setCandidateFilter(event.target.value)}
          >
            <option value="">All candidates</option>
            {candidates.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Visibility</span>
          <select
            className={styles.select}
            value={visibilityFilter}
            onChange={(event) => setVisibilityFilter(event.target.value)}
          >
            <option value="">All</option>
            <option value="yes">Customer visible</option>
            <option value="no">Hidden from customer</option>
          </select>
        </label>
        <motion.button
          type="button"
          className={styles.secondaryButton}
          whileTap={{ scale: 0.97, transition: { duration: 0.15 } }}
          onClick={() => void load()}
        >
          Refresh
        </motion.button>
      </div>

      {loading ? (
        <LoadingState label="Refreshing documents…" />
      ) : (
        <DocumentsBrowseView
          rows={filteredRows}
          path={path}
          direction={direction}
          busyId={busyId}
          onNavigate={navigate}
          onEditMetadata={openEdit}
          onSetVisibility={setVisibility}
        />
      )}

      <AdminDrawer
        open={drawerOpen}
        title="Edit Documents"
        onClose={() => setDrawerOpen(false)}
        footer={
          <>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setDrawerOpen(false)}
            >
              Cancel
            </button>
            <motion.button
              type="button"
              className={styles.primaryButton}
              disabled={saving}
              whileTap={{ scale: 0.97, transition: { duration: 0.15 } }}
              onClick={() => void persist()}
            >
              {saving ? "Saving…" : "Save"}
            </motion.button>
          </>
        }
      >
        {formError ? <p className={styles.formError}>{formError}</p> : null}
        <div className={styles.formGrid}>
          {fields.map((field) => {
            if (field.type === "boolean") {
              return (
                <label key={field.name} className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={Boolean(form[field.name])}
                    onChange={(event) =>
                      setForm((current) => ({
                        ...current,
                        [field.name]: event.target.checked,
                      }))
                    }
                  />
                  {field.label}
                </label>
              );
            }

            if (field.type === "company") {
              return (
                <label key={field.name} className={styles.field}>
                  <span className={styles.fieldLabel}>{field.label}</span>
                  <select
                    className={styles.select}
                    value={String(form[field.name] ?? "")}
                    onChange={(event) =>
                      setForm((current) => ({
                        ...current,
                        [field.name]: event.target.value,
                      }))
                    }
                  >
                    <option value="">Select…</option>
                    {companies.map((company) => (
                      <option key={company.id} value={company.companyName}>
                        {company.companyName}
                      </option>
                    ))}
                  </select>
                </label>
              );
            }

            return (
              <label key={field.name} className={styles.field}>
                <span className={styles.fieldLabel}>{field.label}</span>
                <input
                  className={styles.input}
                  type="text"
                  value={String(form[field.name] ?? "")}
                  onChange={(event) =>
                    setForm((current) => ({
                      ...current,
                      [field.name]: event.target.value,
                    }))
                  }
                />
              </label>
            );
          })}
        </div>
      </AdminDrawer>
    </div>
  );
}

"use client";

import { useCallback, useEffect, useMemo, useState, type ReactNode } from "react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";

import { AdminConfirm } from "@/components/admin/AdminConfirm";
import { AdminDrawer } from "@/components/admin/AdminDrawer";
import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import type { Company } from "@/types/models";

import styles from "./admin.module.css";

export type AdminFieldType =
  | "text"
  | "email"
  | "date"
  | "datetime"
  | "textarea"
  | "select"
  | "boolean"
  | "company";

export interface AdminFieldConfig {
  name: string;
  label: string;
  type: AdminFieldType;
  required?: boolean;
  readOnly?: boolean;
  options?: Array<{ value: string; label: string }>;
  placeholder?: string;
  /** Optional section heading shown above this field group in the drawer. */
  section?: string;
}

export interface AdminColumn<T> {
  key: string;
  header: string;
  render: (row: T) => ReactNode;
}

interface AdminCrudPageProps<T extends { id: string }> {
  title: string;
  description: string;
  columns: AdminColumn<T>[];
  fields: AdminFieldConfig[];
  listUrl: string;
  createUrl?: string;
  updateUrl: (id: string) => string;
  /** When set, shows per-row Delete. */
  deleteUrl?: (id: string) => string;
  /** When set with enableBulkDelete, posts `{ ids: string[] }`. */
  bulkDeleteUrl?: string;
  enableBulkDelete?: boolean;
  /** Extra confirm text for destructive deletes (e.g. company cascade). */
  deleteConfirmExtra?: string;
  initialRows: T[];
  mapResponse: (payload: unknown) => T[];
  companies?: Company[];
  enableCompanyFilter?: boolean;
  getCompanyName?: (row: T) => string | null | undefined;
  searchKeys?: Array<(row: T) => string | null | undefined>;
  rowClassName?: (row: T) => string | undefined;
  rowFilter?: (row: T) => boolean;
  emptyLabel?: string;
  allowCreate?: boolean;
  confirmInactive?: boolean;
  drawerWide?: boolean;
  editLabel?: string;
  toolbarExtra?: ReactNode;
  /** Schema / data-quality warnings shown above the toolbar. */
  warnings?: string[];
  extraActions?: (
    row: T,
    helpers: { reload: () => Promise<void> },
  ) => ReactNode;
  breadcrumbs?: Array<{ label: string; href?: string }>;
  /** Horizontal scroll for many columns (Company List / Excel-width tables). */
  wideTable?: boolean;
}

type FormState = Record<string, string | boolean>;

function toFormValue(value: unknown, type: AdminFieldType): string | boolean {
  if (typeof value === "boolean") return value;
  if (value === null || value === undefined) return "";
  if (type === "date") {
    return String(value).slice(0, 10);
  }
  if (type === "datetime") {
    const date = new Date(String(value));
    if (Number.isNaN(date.getTime())) {
      return String(value).slice(0, 16);
    }
    const pad = (n: number) => String(n).padStart(2, "0");
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }
  return String(value);
}

function buildInitialForm(
  fields: AdminFieldConfig[],
  row?: Record<string, unknown>,
): FormState {
  const state: FormState = {};
  for (const field of fields) {
    if (row && field.name in row) {
      state[field.name] = toFormValue(row[field.name], field.type);
    } else if (field.type === "boolean") {
      state[field.name] =
        field.name === "customerVisible" || field.name === "canView";
    } else if (field.name === "status" && field.type === "select") {
      state[field.name] = "Active";
    } else if (field.name === "trainingOutcome") {
      state[field.name] = "Pass";
    } else if (field.name === "roleType") {
      state[field.name] = "Customer";
    } else {
      state[field.name] = "";
    }
  }
  return state;
}

async function readError(response: Response): Promise<string> {
  return readPublicApiError(response);
}

export function AdminCrudPage<T extends { id: string }>({
  title,
  description,
  columns,
  fields,
  listUrl,
  createUrl,
  updateUrl,
  deleteUrl,
  bulkDeleteUrl,
  enableBulkDelete = false,
  deleteConfirmExtra,
  initialRows,
  mapResponse,
  companies = [],
  enableCompanyFilter = false,
  getCompanyName,
  searchKeys,
  rowClassName,
  rowFilter,
  emptyLabel = "No records found.",
  allowCreate = true,
  confirmInactive = true,
  drawerWide = false,
  editLabel = "Edit",
  toolbarExtra,
  warnings = [],
  extraActions,
  breadcrumbs,
  wideTable = false,
}: AdminCrudPageProps<T>) {
  const { pushToast } = useAdminToast();
  const router = useRouter();
  const pathname = usePathname();
  const searchParams = useSearchParams();
  const [rows, setRows] = useState<T[]>(initialRows);
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState("");
  const [companyFilter, setCompanyFilter] = useState("");
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editing, setEditing] = useState<T | null>(null);
  const [form, setForm] = useState<FormState>({});
  const [formError, setFormError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [pendingSave, setPendingSave] = useState<FormState | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [deleting, setDeleting] = useState(false);

  function openCreate() {
    setEditing(null);
    setForm(buildInitialForm(fields));
    setFormError(null);
    setDrawerOpen(true);
  }

  useEffect(() => {
    const action = searchParams.get("action");
    if (
      allowCreate &&
      (action === "add" || action === "create" || action === "upload")
    ) {
      openCreate();
      router.replace(pathname);
    }
    // Intentionally run once when action is present in the URL.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [searchParams]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const response = await fetch(listUrl, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(await readError(response));
      }
      const payload = await response.json();
      setRows(mapResponse(payload));
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to load records",
        "error",
      );
    } finally {
      setLoading(false);
    }
  }, [listUrl, mapResponse, pushToast]);

  const filtered = useMemo(() => {
    let next = rows;
    if (rowFilter) {
      next = next.filter(rowFilter);
    }
    if (enableCompanyFilter && companyFilter && getCompanyName) {
      next = next.filter(
        (row) =>
          (getCompanyName(row) ?? "").trim().toLowerCase() ===
          companyFilter.trim().toLowerCase(),
      );
    }

    const query = search.trim().toLowerCase();
    if (!query) return next;

    return next.filter((row) => {
      const values = (searchKeys ?? [
        () =>
          Object.values(row as unknown as Record<string, unknown>)
            .filter((value) => typeof value === "string")
            .join(" "),
      ]).map((getter) => getter(row));
      return values.filter(Boolean).join(" ").toLowerCase().includes(query);
    });
  }, [
    rows,
    search,
    searchKeys,
    enableCompanyFilter,
    companyFilter,
    getCompanyName,
    rowFilter,
  ]);

  const allFilteredSelected =
    filtered.length > 0 && filtered.every((row) => selectedIds.has(row.id));

  function toggleSelect(id: string) {
    setSelectedIds((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAllFiltered() {
    if (allFilteredSelected) {
      setSelectedIds(new Set());
      return;
    }
    setSelectedIds(new Set(filtered.map((row) => row.id)));
  }

  async function deleteOne(id: string) {
    if (!deleteUrl) return;
    const ok = window.confirm(
      `Delete this record?${deleteConfirmExtra ? `\n\n${deleteConfirmExtra}` : "\n\nThis cannot be undone."}`,
    );
    if (!ok) return;
    setDeleting(true);
    try {
      const response = await fetch(deleteUrl(id), { method: "DELETE" });
      if (!response.ok) throw new Error(await readError(response));
      pushToast("Record deleted", "success");
      setSelectedIds((current) => {
        const next = new Set(current);
        next.delete(id);
        return next;
      });
      await load();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Delete failed",
        "error",
      );
    } finally {
      setDeleting(false);
    }
  }

  async function deleteSelected() {
    if (!enableBulkDelete || !bulkDeleteUrl || selectedIds.size === 0) return;
    const ids = Array.from(selectedIds);
    const ok = window.confirm(
      `Delete ${ids.length} selected record(s)?${
        deleteConfirmExtra
          ? `\n\n${deleteConfirmExtra}`
          : "\n\nThis cannot be undone."
      }`,
    );
    if (!ok) return;
    setDeleting(true);
    try {
      const response = await fetch(bulkDeleteUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ids }),
      });
      if (!response.ok) throw new Error(await readError(response));
      const payload = (await response.json()) as {
        companiesDeleted?: number;
        relatedDeleted?: number;
      };
      pushToast(
        `Deleted ${payload.companiesDeleted ?? ids.length} record(s)` +
          (payload.relatedDeleted
            ? ` and ${payload.relatedDeleted} related item(s)`
            : ""),
        "success",
      );
      setSelectedIds(new Set());
      await load();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Bulk delete failed",
        "error",
      );
    } finally {
      setDeleting(false);
    }
  }

  function openEdit(row: T) {
    setEditing(row);
    setForm(
      buildInitialForm(fields, row as unknown as Record<string, unknown>),
    );
    setFormError(null);
    setDrawerOpen(true);
  }

  function validate(next: FormState): string | null {
    for (const field of fields) {
      const value = next[field.name];
      if (field.required && field.type !== "boolean") {
        if (value === "" || value === null || value === undefined) {
          return `${field.label} is required.`;
        }
      }
      if (
        field.type === "email" &&
        typeof value === "string" &&
        value.trim() &&
        !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim())
      ) {
        return `${field.label} must be a valid email address.`;
      }
    }
    return null;
  }

  async function persist(next: FormState) {
    setSaving(true);
    setFormError(null);
    try {
      const body: Record<string, unknown> = { ...next };
      if (body.companyId && companies.length) {
        const company = companies.find((item) => item.id === body.companyId);
        if (company) {
          body.companyName = company.companyName;
        }
      }

      const isCreate = !editing;
      const response = await fetch(
        isCreate ? (createUrl ?? listUrl) : updateUrl(editing.id),
        {
          method: isCreate ? "POST" : "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        },
      );

      if (!response.ok) {
        throw new Error(await readError(response));
      }

      const payload = (await response.json().catch(() => null)) as {
        warning?: string;
      } | null;
      pushToast(isCreate ? "Record created." : "Record updated.");
      if (payload?.warning?.trim()) {
        pushToast(payload.warning, "error");
      }
      setDrawerOpen(false);
      await load();
    } catch (error) {
      const message = error instanceof Error ? error.message : "Save failed.";
      setFormError(message);
      pushToast(message, "error");
    } finally {
      setSaving(false);
      setPendingSave(null);
      setConfirmOpen(false);
    }
  }

  function requestSave() {
    const error = validate(form);
    if (error) {
      setFormError(error);
      return;
    }

    if (
      confirmInactive &&
      typeof form.status === "string" &&
      form.status.toLowerCase() === "inactive"
    ) {
      setPendingSave(form);
      setConfirmOpen(true);
      return;
    }

    void persist(form);
  }

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={
              breadcrumbs ?? [
                { label: "Admin", href: "/admin" },
                { label: title },
              ]
            }
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>{title}</h1>
          <p className={styles.subtitle}>{description}</p>
        </div>
        {allowCreate ? (
          <button
            type="button"
            className={styles.primaryButton}
            onClick={openCreate}
          >
            Add new
          </button>
        ) : null}
      </header>

      {warnings.length > 0 ? (
        <div className={styles.schemaWarnings} role="alert">
          {warnings.map((warning) => (
            <p key={warning}>{warning}</p>
          ))}
        </div>
      ) : null}

      <div className={styles.crudToolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            value={search}
            placeholder="Search records…"
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
        {enableCompanyFilter ? (
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Company</span>
            <select
              className={styles.select}
              value={companyFilter}
              onChange={(event) => setCompanyFilter(event.target.value)}
            >
              <option value="">All companies</option>
              {companies.map((company) => (
                <option key={company.id} value={company.companyName}>
                  {company.companyName}
                </option>
              ))}
            </select>
          </label>
        ) : null}
        {toolbarExtra}
        {enableBulkDelete ? (
          <>
            <button
              type="button"
              className={styles.secondaryButton}
              disabled={filtered.length === 0 || deleting}
              onClick={toggleSelectAllFiltered}
            >
              {allFilteredSelected ? "Clear selection" : "Select all"}
            </button>
            <button
              type="button"
              className={styles.dangerButton}
              disabled={selectedIds.size === 0 || deleting}
              onClick={() => {
                void deleteSelected();
              }}
            >
              {deleting
                ? "Deleting…"
                : `Delete selected (${selectedIds.size})`}
            </button>
          </>
        ) : null}
      </div>

      {loading ? (
        <LoadingState label="Refreshing records…" />
      ) : filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No records</h2>
          <p>{emptyLabel}</p>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table
            className={`${styles.dataTable}${
              wideTable ? ` ${styles.companiesWideTable}` : ""
            }`}
          >
            <thead>
              <tr>
                {enableBulkDelete || deleteUrl ? (
                  <th scope="col" className={styles.selectCol}>
                    {enableBulkDelete ? (
                      <input
                        type="checkbox"
                        checked={allFilteredSelected}
                        onChange={toggleSelectAllFiltered}
                        aria-label="Select all"
                        disabled={deleting}
                      />
                    ) : null}
                  </th>
                ) : null}
                {columns.map((column) => (
                  <th key={column.key} scope="col">
                    {column.header}
                  </th>
                ))}
                <th scope="col">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row) => (
                <tr key={row.id} className={rowClassName?.(row)}>
                  {enableBulkDelete || deleteUrl ? (
                    <td className={styles.selectCol}>
                      {enableBulkDelete ? (
                        <input
                          type="checkbox"
                          checked={selectedIds.has(row.id)}
                          onChange={() => toggleSelect(row.id)}
                          aria-label={`Select ${row.id}`}
                          disabled={deleting}
                        />
                      ) : null}
                    </td>
                  ) : null}
                  {columns.map((column) => (
                    <td key={column.key}>{column.render(row)}</td>
                  ))}
                  <td>
                    <button
                      type="button"
                      className={styles.linkButton}
                      onClick={() => openEdit(row)}
                    >
                      {editLabel}
                    </button>
                    {deleteUrl ? (
                      <>
                        {" · "}
                        <button
                          type="button"
                          className={styles.linkButtonDanger}
                          disabled={deleting}
                          onClick={() => {
                            void deleteOne(row.id);
                          }}
                        >
                          Delete
                        </button>
                      </>
                    ) : null}
                    {extraActions ? (
                      <> · {extraActions(row, { reload: load })}</>
                    ) : null}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <AdminDrawer
        open={drawerOpen}
        title={editing ? `Edit ${title}` : `Add ${title}`}
        onClose={() => setDrawerOpen(false)}
        wide={drawerWide}
        footer={
          <>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setDrawerOpen(false)}
            >
              Cancel
            </button>
            <button
              type="button"
              className={styles.primaryButton}
              disabled={saving}
              onClick={requestSave}
            >
              {saving ? "Saving…" : "Save"}
            </button>
          </>
        }
      >
        {formError ? <p className={styles.formError}>{formError}</p> : null}
        <div className={styles.formGrid}>
          {fields.map((field, index) => {
            const previousSection =
              index > 0 ? fields[index - 1]?.section : undefined;
            const showSection =
              Boolean(field.section) && field.section !== previousSection;

            let control: ReactNode;
            if (field.type === "boolean") {
              control = (
                <label className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={Boolean(form[field.name])}
                    disabled={field.readOnly}
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
            } else if (field.type === "textarea") {
              control = (
                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span className={styles.fieldLabel}>{field.label}</span>
                  <textarea
                    className={styles.input}
                    rows={4}
                    value={String(form[field.name] ?? "")}
                    placeholder={field.placeholder}
                    readOnly={field.readOnly}
                    disabled={field.readOnly}
                    onChange={(event) =>
                      setForm((current) => ({
                        ...current,
                        [field.name]: event.target.value,
                      }))
                    }
                  />
                </label>
              );
            } else if (field.type === "company" || field.type === "select") {
              const options =
                field.type === "company"
                  ? companies.map((company) => ({
                      value:
                        field.name === "companyId"
                          ? company.id
                          : company.companyName,
                      label: company.companyName,
                    }))
                  : (field.options ?? []);

              control = (
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>{field.label}</span>
                  <select
                    className={styles.select}
                    value={String(form[field.name] ?? "")}
                    disabled={field.readOnly}
                    onChange={(event) =>
                      setForm((current) => ({
                        ...current,
                        [field.name]: event.target.value,
                      }))
                    }
                  >
                    <option value="">Select…</option>
                    {options.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
              );
            } else {
              control = (
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>{field.label}</span>
                  <input
                    className={styles.input}
                    type={
                      field.type === "date"
                        ? "date"
                        : field.type === "datetime"
                          ? "datetime-local"
                          : field.type === "email"
                            ? "email"
                            : "text"
                    }
                    value={String(form[field.name] ?? "")}
                    placeholder={field.placeholder}
                    readOnly={field.readOnly}
                    disabled={field.readOnly}
                    onChange={(event) =>
                      setForm((current) => ({
                        ...current,
                        [field.name]: event.target.value,
                      }))
                    }
                  />
                </label>
              );
            }

            return (
              <div key={field.name} className={styles.formFieldBlock}>
                {showSection ? (
                  <h3 className={styles.formSectionTitle}>{field.section}</h3>
                ) : null}
                {control}
              </div>
            );
          })}
        </div>
      </AdminDrawer>

      <AdminConfirm
        open={confirmOpen}
        title="Confirm status change"
        message="You are setting this record to Inactive. Continue?"
        confirmLabel="Set Inactive"
        onCancel={() => {
          setConfirmOpen(false);
          setPendingSave(null);
        }}
        onConfirm={() => {
          if (pendingSave) {
            void persist(pendingSave);
          }
        }}
      />
    </div>
  );
}

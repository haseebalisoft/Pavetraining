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
  | "multiselect"
  | "boolean"
  | "company"
  | "workforce";

export interface AdminWorkforceOption {
  id: string;
  candidateName: string;
  companyName: string;
  companyId?: string | null;
  nporsNumbers?: string | null;
  eusrNumber?: string | null;
  swqrNumber?: string | null;
  inHouseCertificationNumber?: string | null;
  workforceNumber?: string | null;
}

/** Minimal Permissions row for company-scoped TM/Supervisor selects. */
export interface AdminPermissionPersonOption {
  id: string;
  userEmail: string;
  name: string | null;
  status: string;
  permissionRole: "Admin" | "Customer" | "Candidate";
  /** Live SharePoint RoleType (Training Manager / Supervisor / …). */
  sharePointRoleType?: string | null;
  companyId: string | null;
  companyName: string | null;
}

/** Department rows for company-scoped Workforce / Permissions coverage. */
export interface AdminDepartmentOption {
  id: string;
  name: string;
  companyId: string | null;
  companyName: string | null;
}

export interface AdminFieldConfig {
  name: string;
  label: string;
  type: AdminFieldType;
  required?: boolean;
  readOnly?: boolean;
  options?: Array<{ value: string; label: string }>;
  /**
   * When set, select options are built from `permissionPeople` filtered to the
   * form's selected company (and this Permissions form role).
   */
  permissionRoleFilter?: "Admin" | "Customer" | "Candidate";
  /**
   * When true, select/multiselect options come from `departments` filtered to
   * the form's selected company. Values are department ids (multiselect) or
   * names (select) depending on `departmentValueMode`.
   */
  companyScopedDepartments?: boolean;
  /** How to store the selected department in the form (default: name). */
  departmentValueMode?: "id" | "name";
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
  updateUrl?: (id: string) => string;
  /** When set, shows per-row Delete. */
  deleteUrl?: (id: string) => string;
  /** When set with enableBulkDelete, posts `{ ids: string[] }`. */
  bulkDeleteUrl?: string;
  enableBulkDelete?: boolean;
  /** Extra confirm text for destructive deletes (e.g. company cascade). */
  deleteConfirmExtra?: string;
  initialRows: T[];
  mapResponse: (payload: unknown) => T[];
  /** Receives refreshed rows so alternate views can share the same data. */
  onRowsChange?: (rows: T[]) => void;
  companies?: Company[];
  /** Workforce options for type="workforce" fields (auto-fills name/company). */
  workforce?: AdminWorkforceOption[];
  /** Permissions people for company-scoped TM / Supervisor selects. */
  permissionPeople?: AdminPermissionPersonOption[];
  /** Departments for company-scoped department select / coverage multiselect. */
  departments?: AdminDepartmentOption[];
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
  /** Freeze first two data columns while scrolling horizontally. */
  stickyLeadColumns?: boolean;
  /**
   * Pin a single identity column (e.g. Company Name / Candidate Name) to the
   * left edge while scrolling horizontally. Pass the matching `column.key`.
   */
  stickyColumnKey?: string;
  /** Extra table class (e.g. matrix grid layout). */
  tableClassName?: string;
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

function matchWorkforceId(
  form: FormState,
  workforce: AdminWorkforceOption[],
): string {
  const name = String(form.candidateName ?? "").trim().toLowerCase();
  const company = String(form.companyName ?? "").trim().toLowerCase();
  if (!name) return "";
  const exact =
    workforce.find(
      (row) =>
        row.candidateName.trim().toLowerCase() === name &&
        row.companyName.trim().toLowerCase() === company,
    ) ??
    workforce.find(
      (row) => row.candidateName.trim().toLowerCase() === name,
    );
  return exact?.id ?? "";
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
  onRowsChange,
  companies = [],
  workforce = [],
  permissionPeople = [],
  departments = [],
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
  stickyLeadColumns = false,
  stickyColumnKey,
  tableClassName,
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
  const [workforceQuery, setWorkforceQuery] = useState("");
  const [saving, setSaving] = useState(false);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [pendingSave, setPendingSave] = useState<FormState | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [deleting, setDeleting] = useState(false);
  const [livePermissionPeople, setLivePermissionPeople] =
    useState<AdminPermissionPersonOption[]>(permissionPeople);
  const needsPermissionPeople = fields.some(
    (field) => Boolean(field.permissionRoleFilter),
  );

  useEffect(() => {
    setLivePermissionPeople(permissionPeople);
  }, [permissionPeople]);

  const refreshPermissionPeople = useCallback(async () => {
    if (!needsPermissionPeople) return;
    try {
      const response = await fetch("/api/admin/permissions", {
        cache: "no-store",
      });
      if (!response.ok) return;
      const payload = (await response.json().catch(() => null)) as {
        records?: AdminPermissionPersonOption[];
      } | null;
      const records = payload?.records ?? [];
      setLivePermissionPeople(
        records.map((row) => ({
          id: row.id,
          userEmail: row.userEmail,
          name: row.name ?? null,
          status: row.status,
          permissionRole: row.permissionRole,
          sharePointRoleType:
            (row as { sharePointRoleType?: string | null }).sharePointRoleType ??
            null,
          companyId: row.companyId ?? null,
          companyName: row.companyName ?? null,
        })),
      );
    } catch {
      // Keep last good list if SharePoint refresh fails.
    }
  }, [needsPermissionPeople]);

  function openCreate() {
    setEditing(null);
    setForm(buildInitialForm(fields));
    setFormError(null);
    setWorkforceQuery("");
    void refreshPermissionPeople();
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
      const nextRows = mapResponse(payload);
      setRows(nextRows);
      onRowsChange?.(nextRows);
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to load records",
        "error",
      );
    } finally {
      setLoading(false);
    }
  }, [listUrl, mapResponse, onRowsChange, pushToast]);

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
    setWorkforceQuery("");
    void refreshPermissionPeople();
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
      const workforceId = matchWorkforceId(next, workforce);
      if (workforceId) {
        body.workforceId = workforceId;
      }

      const isCreate = !editing;
      if (!isCreate && !updateUrl) {
        throw new Error("Editing is not available for this list.");
      }
      const response = await fetch(
        isCreate ? (createUrl ?? listUrl) : updateUrl!(editing.id),
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
        choiceWarnings?: string[];
        matrixSeedWarning?: string;
        matrixSync?: {
          summary?: {
            updated?: number;
            created?: number;
            skipped?: number;
            errors?: number;
          };
        };
      } | null;
      const sync = payload?.matrixSync?.summary;
      const syncNote = sync
        ? ` Matrix sync: ${sync.updated ?? 0} updated, ${sync.created ?? 0} created` +
          (sync.errors ? `, ${sync.errors} error(s)` : "") +
          (sync.skipped ? `, ${sync.skipped} skipped` : "") +
          "."
        : "";
      pushToast(
        (isCreate ? "Record created." : "Record updated.") + syncNote,
      );
      const warnings = [
        payload?.warning?.trim(),
        payload?.matrixSeedWarning?.trim(),
        ...(payload?.choiceWarnings ?? []).map((part) => part.trim()),
      ].filter(Boolean) as string[];
      for (const warning of warnings) {
        pushToast(warning, "error");
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
        <div
          className={`${styles.tableWrap}${
            wideTable ? ` ${styles.tableWrapWide}` : ""
          }`}
        >
          <table
            className={[
              styles.dataTable,
              wideTable ? styles.companiesWideTable : "",
              stickyLeadColumns ? styles.stickyLeadColumns : "",
              tableClassName || "",
            ]
              .filter(Boolean)
              .join(" ")}
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
                  <th
                    key={column.key}
                    scope="col"
                    className={
                      stickyColumnKey && column.key === stickyColumnKey
                        ? styles.stickyIdentityCell
                        : undefined
                    }
                  >
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
                    <td
                      key={column.key}
                      className={
                        stickyColumnKey && column.key === stickyColumnKey
                          ? styles.stickyIdentityCell
                          : undefined
                      }
                    >
                      {column.render(row)}
                    </td>
                  ))}
                  <td>
                    {updateUrl ? (
                      <button
                        type="button"
                        className={styles.linkButton}
                        onClick={() => openEdit(row)}
                      >
                        {editLabel}
                      </button>
                    ) : null}
                    {deleteUrl ? (
                      <>
                        {updateUrl ? " · " : null}
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
                      <>
                        {updateUrl || deleteUrl ? " · " : null}
                        {extraActions(row, { reload: load })}
                      </>
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
            } else if (field.type === "workforce") {
              const selectedId = matchWorkforceId(form, workforce);
              const selectedCompanyName = String(form.companyName ?? "")
                .trim()
                .toLowerCase();
              const selectedCompanyId = String(form.companyId ?? "").trim();
              const resolvedCompanyId =
                selectedCompanyId ||
                companies.find(
                  (company) =>
                    company.companyName.trim().toLowerCase() ===
                    selectedCompanyName,
                )?.id ||
                "";
              const hasCompany = Boolean(
                selectedCompanyName || resolvedCompanyId,
              );
              const query = workforceQuery.trim().toLowerCase();
              const companyWorkforce = hasCompany
                ? workforce.filter((row) => {
                    if (
                      resolvedCompanyId &&
                      row.companyId &&
                      row.companyId === resolvedCompanyId
                    ) {
                      return true;
                    }
                    return (
                      Boolean(selectedCompanyName) &&
                      row.companyName.trim().toLowerCase() ===
                        selectedCompanyName
                    );
                  })
                : [];
              const filteredWorkforce = query
                ? companyWorkforce.filter((row) => {
                    const haystack = [
                      row.candidateName,
                      row.companyName,
                      row.workforceNumber,
                      row.nporsNumbers,
                      row.eusrNumber,
                      row.swqrNumber,
                      row.inHouseCertificationNumber,
                    ]
                      .filter(Boolean)
                      .join(" ")
                      .toLowerCase();
                    return haystack.includes(query);
                  })
                : companyWorkforce;
              const candidateDisabled = field.readOnly || !hasCompany;
              const emptyOptionLabel = !hasCompany
                ? "Select a company first…"
                : companyWorkforce.length === 0
                  ? "No Workforce candidates for this company"
                  : query && filteredWorkforce.length === 0
                    ? `No matches for “${workforceQuery.trim()}”`
                    : "Select candidate…";
              control = (
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {field.label}
                    {hasCompany
                      ? " (scoped to selected company)"
                      : " — select a company first"}
                  </span>
                  <input
                    className={`${styles.input} ${styles.workforceFilter}`}
                    type="search"
                    value={workforceQuery}
                    placeholder={
                      hasCompany
                        ? "Search name or workforce / cert number…"
                        : "Select a company first…"
                    }
                    disabled={candidateDisabled}
                    onChange={(event) => setWorkforceQuery(event.target.value)}
                  />
                  <select
                    className={styles.select}
                    value={selectedId}
                    disabled={candidateDisabled}
                    onChange={(event) => {
                      const hit = workforce.find(
                        (row) => row.id === event.target.value,
                      );
                      if (!hit) {
                        setForm((current) => ({
                          ...current,
                          candidateName: "",
                          companyName: current.companyName,
                          nporsNumber: "",
                          eusrNumber: "",
                          swqrNumber: "",
                          inHouseCertificationNumber: "",
                        }));
                        return;
                      }
                      setForm((current) => ({
                        ...current,
                        candidateName: hit.candidateName,
                        companyName: hit.companyName,
                        // Always refresh projected Workforce numbers for the chosen candidate.
                        nporsNumber: hit.nporsNumbers || "",
                        eusrNumber: hit.eusrNumber || "",
                        swqrNumber: hit.swqrNumber || "",
                        inHouseCertificationNumber:
                          hit.inHouseCertificationNumber || "",
                        workforceNumber: hit.workforceNumber || "",
                      }));
                    }}
                  >
                    <option value="">{emptyOptionLabel}</option>
                    {filteredWorkforce.map((row) => (
                      <option key={row.id} value={row.id}>
                        {row.workforceNumber?.trim()
                          ? `${row.candidateName} (${row.workforceNumber})`
                          : row.candidateName}
                      </option>
                    ))}
                  </select>
                  {hasCompany && companyWorkforce.length === 0 ? (
                    <p className={styles.helpText}>
                      Add this person under Admin → Workforce for this company
                      first.
                    </p>
                  ) : null}
                  {hasCompany &&
                  query &&
                  companyWorkforce.length > 0 &&
                  filteredWorkforce.length === 0 ? (
                    <p className={styles.helpText}>
                      Clear the search box to see all{" "}
                      {companyWorkforce.length} candidate
                      {companyWorkforce.length === 1 ? "" : "s"} for this
                      company.
                    </p>
                  ) : null}
                </label>
              );
            } else if (field.type === "multiselect") {
              const selected = String(form[field.name] ?? "")
                .split(/[;,|]+/)
                .map((part) => part.trim())
                .filter(Boolean);
              const selectedSet = new Set(
                selected.map((part) => part.toLowerCase()),
              );
              const selectedCompanyName = String(form.companyName ?? "")
                .trim()
                .toLowerCase();
              const selectedCompanyId = String(form.companyId ?? "").trim();
              const valueMode = field.departmentValueMode ?? "id";
              let multiOptions = field.options ?? [];
              if (field.companyScopedDepartments) {
                if (!selectedCompanyName && !selectedCompanyId) {
                  multiOptions = [];
                } else {
                  multiOptions = departments
                    .filter((dept) => {
                      if (
                        selectedCompanyId &&
                        dept.companyId &&
                        dept.companyId === selectedCompanyId
                      ) {
                        return true;
                      }
                      const deptCompany = (dept.companyName || "")
                        .trim()
                        .toLowerCase();
                      return (
                        Boolean(selectedCompanyName) &&
                        deptCompany === selectedCompanyName
                      );
                    })
                    .map((dept) => ({
                      value: valueMode === "name" ? dept.name : dept.id,
                      label: dept.name,
                    }));
                }
              }
              control = (
                <fieldset className={styles.field}>
                  <legend className={styles.fieldLabel}>{field.label}</legend>
                  {field.companyScopedDepartments &&
                  !selectedCompanyName &&
                  !selectedCompanyId ? (
                    <p className={styles.helpText}>
                      Select a company first to choose departments.
                    </p>
                  ) : null}
                  {field.companyScopedDepartments &&
                  (selectedCompanyName || selectedCompanyId) &&
                  multiOptions.length === 0 ? (
                    <p className={styles.helpText}>
                      No departments for this company yet. Add them under
                      Departments.
                    </p>
                  ) : null}
                  <div className={styles.multiSelectList}>
                    {multiOptions.map((option) => {
                      const checked =
                        selectedSet.has(option.value.trim().toLowerCase()) ||
                        selectedSet.has(option.label.trim().toLowerCase());
                      return (
                        <label
                          key={option.value}
                          className={styles.multiSelectOption}
                        >
                          <input
                            type="checkbox"
                            checked={checked}
                            disabled={field.readOnly}
                            onChange={() => {
                              setForm((current) => {
                                const currentParts = String(
                                  current[field.name] ?? "",
                                )
                                  .split(/[;,|]+/)
                                  .map((part) => part.trim())
                                  .filter(Boolean);
                                const exists = currentParts.some(
                                  (part) =>
                                    part.toLowerCase() ===
                                    option.value.trim().toLowerCase(),
                                );
                                const nextParts = exists
                                  ? currentParts.filter(
                                      (part) =>
                                        part.toLowerCase() !==
                                        option.value.trim().toLowerCase(),
                                    )
                                  : [...currentParts, option.value];
                                return {
                                  ...current,
                                  [field.name]: nextParts.join(", "),
                                };
                              });
                            }}
                          />
                          <span>{option.label}</span>
                        </label>
                      );
                    })}
                  </div>
                </fieldset>
              );
            } else if (field.type === "company" || field.type === "select") {
              const selectedCompanyName = String(form.companyName ?? "")
                .trim()
                .toLowerCase();
              const selectedCompanyId = String(form.companyId ?? "").trim();

              let options: Array<{ value: string; label: string }>;
              if (field.type === "company") {
                options = companies.map((company) => ({
                  value:
                    field.name === "companyId"
                      ? company.id
                      : company.companyName,
                  label: company.companyName,
                }));
              } else if (field.permissionRoleFilter) {
                const role = field.permissionRoleFilter;
                const emptyLabel =
                  role === "Admin"
                    ? "No active Training Managers in Permissions for this company"
                    : role === "Customer"
                      ? "No active Supervisors in Permissions for this company"
                      : "No people assigned to this company yet";
                const resolvedCompanyId =
                  selectedCompanyId ||
                  companies.find(
                    (company) =>
                      company.companyName.trim().toLowerCase() ===
                      selectedCompanyName,
                  )?.id ||
                  "";
                if (!selectedCompanyName && !resolvedCompanyId) {
                  options = [
                    { value: "", label: "Select a company first…" },
                  ];
                } else {
                  const personMatchesCompany = (person: {
                    companyId: string | null;
                    companyName: string | null;
                  }) => {
                    if (
                      resolvedCompanyId &&
                      person.companyId &&
                      person.companyId === resolvedCompanyId
                    ) {
                      return true;
                    }
                    const personCompany = (person.companyName || "")
                      .trim()
                      .toLowerCase();
                    return (
                      Boolean(selectedCompanyName) &&
                      personCompany === selectedCompanyName
                    );
                  };

                  const matchesRoleType = (
                    person: AdminPermissionPersonOption,
                  ) => {
                    const spRole = (person.sharePointRoleType || "")
                      .trim()
                      .toLowerCase()
                      .replace(/\s+/g, " ");
                    if (role === "Admin") {
                      // Strict: SharePoint RoleType = Training Manager only.
                      return (
                        spRole === "training manager" ||
                        spRole === "trainingmanager" ||
                        (!spRole && person.permissionRole === "Admin")
                      );
                    }
                    if (role === "Customer") {
                      return (
                        spRole === "supervisor" ||
                        (!spRole && person.permissionRole === "Customer")
                      );
                    }
                    return person.permissionRole === role;
                  };

                  const scoped = livePermissionPeople.filter((person) => {
                    if ((person.status || "").toLowerCase() !== "active") {
                      return false;
                    }
                    if (!matchesRoleType(person)) return false;
                    return personMatchesCompany(person);
                  });

                  const peopleOptions = scoped
                    .map((person) => {
                      const value = person.name?.trim() || person.userEmail;
                      if (!value) return null;
                      return {
                        value,
                        label: person.name?.trim()
                          ? `${person.name.trim()} (${person.userEmail})`
                          : person.userEmail,
                      };
                    })
                    .filter(
                      (option): option is { value: string; label: string } =>
                        option !== null,
                    )
                    .sort((a, b) => a.label.localeCompare(b.label));
                  options =
                    peopleOptions.length === 0
                      ? [{ value: "", label: emptyLabel }]
                      : [{ value: "", label: "— None —" }, ...peopleOptions];
                }
              } else if (field.companyScopedDepartments) {
                const valueMode = field.departmentValueMode ?? "name";
                if (!selectedCompanyName && !selectedCompanyId) {
                  options = [
                    { value: "", label: "Select a company first…" },
                  ];
                } else {
                  const scoped = departments.filter((dept) => {
                    if (
                      selectedCompanyId &&
                      dept.companyId &&
                      dept.companyId === selectedCompanyId
                    ) {
                      return true;
                    }
                    const deptCompany = (dept.companyName || "")
                      .trim()
                      .toLowerCase();
                    return (
                      Boolean(selectedCompanyName) &&
                      deptCompany === selectedCompanyName
                    );
                  });
                  options =
                    scoped.length === 0
                      ? [
                          {
                            value: "",
                            label:
                              "No departments for this company — add under Departments",
                          },
                        ]
                      : [
                          { value: "", label: "— None —" },
                          ...scoped.map((dept) => ({
                            value:
                              valueMode === "id" ? dept.id : dept.name,
                            label: dept.name,
                          })),
                        ];
                  const currentDept = String(form[field.name] ?? "").trim();
                  if (
                    currentDept &&
                    !options.some(
                      (option) =>
                        option.value.trim().toLowerCase() ===
                        currentDept.toLowerCase(),
                    )
                  ) {
                    options = [
                      ...options,
                      {
                        value: currentDept,
                        label: `${currentDept} (not in Departments list)`,
                      },
                    ];
                  }
                }
              } else {
                options = field.options ?? [];
              }

              control = (
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>{field.label}</span>
                  <select
                    className={styles.select}
                    value={String(form[field.name] ?? "")}
                    disabled={field.readOnly}
                    onChange={(event) => {
                      const value = event.target.value;
                      if (
                        field.type === "company" &&
                        (field.name === "companyName" ||
                          field.name === "companyId")
                      ) {
                        setWorkforceQuery("");
                        void refreshPermissionPeople();
                      }
                      setForm((current) => {
                        if (field.type !== "company") {
                          return { ...current, [field.name]: value };
                        }

                        const next: FormState = {
                          ...current,
                          [field.name]: value,
                        };

                        // Clear company-scoped people / departments when company changes.
                        if (
                          field.name === "companyName" ||
                          field.name === "companyId"
                        ) {
                          next.trainingManager = "";
                          next.supervisor = "";
                          next.department = "";
                          next.departmentsAllowed = "";
                        }

                        if (field.name !== "companyName") {
                          return next;
                        }

                        const candidateName = String(
                          current.candidateName ?? "",
                        ).trim().toLowerCase();
                        const currentCompany = String(
                          current.companyName ?? "",
                        ).trim().toLowerCase();
                        const selectedCandidate =
                          workforce.find(
                            (row) =>
                              row.candidateName.trim().toLowerCase() ===
                                candidateName &&
                              row.companyName.trim().toLowerCase() ===
                                currentCompany,
                          ) ??
                          workforce.find(
                            (row) =>
                              row.candidateName.trim().toLowerCase() ===
                              candidateName,
                          );
                        const candidateStillMatches =
                          !selectedCandidate ||
                          !value ||
                          selectedCandidate.companyName.trim().toLowerCase() ===
                            value.trim().toLowerCase();

                        if (!candidateStillMatches) {
                          next.candidateName = "";
                          next.nporsNumber = "";
                          next.eusrNumber = "";
                          next.swqrNumber = "";
                          next.inHouseCertificationNumber = "";
                          next.workforceNumber = "";
                        }

                        return next;
                      });
                    }}
                  >
                    {field.permissionRoleFilter ? null : (
                      <option value="">Select…</option>
                    )}
                    {options.map((option) => (
                      <option key={`${option.value}-${option.label}`} value={option.value}>
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

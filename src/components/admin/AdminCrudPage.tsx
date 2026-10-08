"use client";

import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";

import { AdminConfirm } from "@/components/admin/AdminConfirm";
import { AdminDrawer } from "@/components/admin/AdminDrawer";
import { useAdminToast } from "@/components/admin/AdminToast";
import { QuickAddPermissionPersonModal } from "@/components/admin/QuickAddPermissionPersonModal";
import type { AdminPermissionRecord } from "@/lib/services/adminCrudService";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import { addCalendarYearsIso, defaultPassExpiryIso } from "@/lib/utils/formatDate";
import { isValidEmail } from "@/lib/validation/email";
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
  niNumber?: string | null;
}

/** Minimal Permissions row for company-scoped TM/Supervisor selects. */
export interface AdminPermissionPersonOption {
  id: string;
  userEmail: string;
  name: string | null;
  status: string;
  permissionRole: "Admin" | "Manager" | "Supervisor" | "Customer" | "Candidate";
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
   * form's selected company AND the ROUTING role bucket (Admin / Customer /
   * Candidate). Broader than `sharePointRoleTypeFilter` — "Admin" here matches
   * both `RoleType = Admin` and `RoleType = Training Manager`. Prefer
   * `sharePointRoleTypeFilter` for strict matching on the raw SharePoint
   * `RoleType` value.
   */
  permissionRoleFilter?: "Admin" | "Customer" | "Candidate";
  /**
   * STRICT SharePoint `RoleType` filter for `permissionPeople` options. Use
   * for pages that need to pick a specific SharePoint role (Training Manager
   * / Supervisor / Candidate / Admin) without the routing-bucket permissive
   * match. Also filters to Active-only rows and to the form's selected
   * company. A name already stored on the row (e.g. Workforce Training
   * Manager) is still shown as the current selection even if that person is
   * missing from the filtered Permissions list. Setting this also enables
   * the "+ Add new" quick-create button on the field (unless
   * `allowQuickAdd={false}` overrides).
   */
  sharePointRoleTypeFilter?:
    | "Admin"
    | "Training Manager"
    | "Supervisor"
    | "Candidate";
  /** Explicit override for the quick-add button next to the dropdown. */
  allowQuickAdd?: boolean;
  /**
   * Controls how the `departments` prop populates a select/multiselect:
   *   - `true` — filter to departments belonging to the form's selected
   *     company (default cross-company hygiene for e.g. Workforce Department).
   *   - `"all"` — show every active department across every company, labelled
   *     with the company name in parentheses. Use on forms where the admin
   *     legitimately needs to pick departments regardless of the row's Company
   *     (e.g. Permissions "Departments allowed" for cross-company TMs).
   *   - `false` / omitted — use `field.options` instead of `departments`.
   * Values are department ids (multiselect) or names (select) depending on
   * `departmentValueMode`.
   */
  companyScopedDepartments?: boolean | "all";
  /** How to store the selected department in the form (default: name). */
  departmentValueMode?: "id" | "name";
  placeholder?: string;
  /** Optional section heading shown above this field group in the drawer. */
  section?: string;
  /**
   * When set on the `expiry` field, a blank expiry is filled as
   * trainingDate + this many calendar years (NPORS / EUSR = 3, Streetworks = 5).
   */
  defaultExpiryYears?: number;
  /**
   * On create, `type: "workforce"` becomes a company-scoped checkbox list so
   * one training entry can be saved for several candidates at once (NPORS).
   * Edit remains a single select. Selected Workforce ids are stored in
   * `candidateWorkforceIds` (semicolon-separated).
   */
  workforceMultiSelect?: boolean;
}

type AdminSelectOption = { value: string; label: string };

function permissionPersonToOption(
  person: AdminPermissionPersonOption,
): AdminSelectOption | null {
  const value = person.name?.trim() || person.userEmail.trim();
  if (!value) return null;
  return {
    value,
    label: person.name?.trim()
      ? `${person.name.trim()} (${person.userEmail})`
      : person.userEmail,
  };
}

/**
 * Keep a Workforce (or similar) assignment visible in a Permissions-backed
 * dropdown. Other options stay Active + role + company filtered; the stored
 * name is injected when it is not in that list so the select can display it
 * and the admin can still pick someone else.
 */
function withAssignedWorkforcePersonOption(
  peopleOptions: AdminSelectOption[],
  assignedRaw: string,
  allPeople: AdminPermissionPersonOption[],
  emptyLabel: string,
): AdminSelectOption[] {
  const assigned = assignedRaw.trim();
  const assignedKey = assigned.toLowerCase();
  const hasExactValue = peopleOptions.some(
    (option) => option.value.trim().toLowerCase() === assignedKey,
  );

  let merged = peopleOptions;
  if (assigned && !hasExactValue) {
    const match = allPeople.find((person) => {
      const name = (person.name || "").trim().toLowerCase();
      const email = (person.userEmail || "").trim().toLowerCase();
      return name === assignedKey || email === assignedKey;
    });
    const fromMatch = match ? permissionPersonToOption(match) : null;
    merged = [
      {
        value: assigned,
        label: fromMatch?.label ?? assigned,
      },
      ...peopleOptions,
    ];
  }

  if (merged.length === 0) {
    return [{ value: "", label: emptyLabel }];
  }
  return [{ value: "", label: "— None —" }, ...merged];
}

export type AdminColumnRenderContext<T extends { id: string }> = {
  reload: () => void | Promise<void>;
  /** Merge fields into the in-memory row after a successful inline save. */
  patchRow: (id: string, patch: Partial<T>) => void;
};

export interface AdminColumn<T extends { id: string }> {
  key: string;
  header: string;
  render: (row: T, ctx: AdminColumnRenderContext<T>) => ReactNode;
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
  /** Defaults merged into the create form (e.g. next company number). */
  getCreateDefaults?: (
    rows: T[],
  ) => Partial<Record<string, string | boolean>>;
  /**
   * Enables optimistic UI: delete removes the row immediately and the API
   * runs in the background. On failure the row is silently restored and
   * added to a small "N action(s) need review" pill above the table.
   * Update + create likewise apply locally first and revert silently on
   * failure. No red toasts on background failure — the review pill is the
   * only surface.
   */
  optimistic?: boolean;
  /** Short row label for the review pill (fallback: `row.id`). */
  optimisticRowLabel?: (row: T) => string;
}

type FormState = Record<string, string | boolean>;

const EMPTY_COMPANIES: Company[] = [];
const EMPTY_WORKFORCE: AdminWorkforceOption[] = [];
const EMPTY_PERMISSION_PEOPLE: AdminPermissionPersonOption[] = [];
const EMPTY_DEPARTMENTS: AdminDepartmentOption[] = [];
const EMPTY_WARNINGS: string[] = [];

function expiryYearsFromFields(fields: AdminFieldConfig[]): number | null {
  const years = fields.find((field) => field.name === "expiry")?.defaultExpiryYears;
  return typeof years === "number" && years > 0 ? years : null;
}

function autoExpiryIso(
  fields: AdminFieldConfig[],
  trainingDate: string,
): string | null {
  const years = expiryYearsFromFields(fields);
  if (years != null) {
    return addCalendarYearsIso(trainingDate, years);
  }
  const trimmed = trainingDate.trim();
  return trimmed ? defaultPassExpiryIso(trimmed) : null;
}

function shouldAutoFillBlankExpiry(
  fields: AdminFieldConfig[],
  outcome: string,
): boolean {
  if (expiryYearsFromFields(fields) != null) return true;
  return outcome.trim().toLowerCase() === "pass";
}

/** Fill blank expiry from trainingDate + defaultExpiryYears (NPORS/EUSR = 3). */
function withAutoFilledBlankExpiry(
  fields: AdminFieldConfig[],
  state: FormState,
): FormState {
  if (String(state.expiry ?? "").trim()) return state;
  if (
    !shouldAutoFillBlankExpiry(fields, String(state.trainingOutcome ?? ""))
  ) {
    return state;
  }
  const filled = autoExpiryIso(fields, String(state.trainingDate ?? ""));
  if (!filled) return state;
  return { ...state, expiry: filled };
}

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
      let value = toFormValue(row[field.name], field.type);
      // SharePoint AccessScope may be blank or legacy wording — map to form choices.
      if (field.name === "accessScope" && field.type === "select") {
        const raw = String(value ?? "").trim().toLowerCase();
        if (!raw) {
          value = "Full Company";
        } else if (raw.includes("candidate")) {
          value = "Candidate Only";
        } else if (raw.includes("department")) {
          value = "Department Only";
        } else if (
          raw.includes("full") ||
          raw === "company" ||
          raw === "all" ||
          raw.includes("all compan")
        ) {
          value = "Full Company";
        } else if (
          field.options?.length &&
          !field.options.some(
            (option) =>
              option.value.trim().toLowerCase() === raw ||
              option.value === value,
          )
        ) {
          value = "Full Company";
        }
      }
      state[field.name] = value;
    } else if (field.type === "boolean") {
      state[field.name] =
        field.name === "customerVisible" || field.name === "canView";
    } else if (field.name === "status" && field.type === "select") {
      state[field.name] = "Active";
    } else if (field.name === "accessScope" && field.type === "select") {
      state[field.name] = "Full Company";
    } else if (field.name === "permissionRole" && field.type === "select") {
      state[field.name] = "Manager";
    } else if (field.name === "trainingOutcome") {
      state[field.name] = "Pass";
    } else if (field.name === "expiry" && field.type === "date") {
      // NPORS (and any field with defaultExpiryYears) stays blank until a
      // training date is entered, then fills as training date + N years.
      state[field.name] =
        typeof field.defaultExpiryYears === "number"
          ? ""
          : defaultPassExpiryIso();
    } else if (field.name === "roleType") {
      state[field.name] = "Customer";
    } else {
      state[field.name] = "";
    }
  }
  return withAutoFilledBlankExpiry(fields, state);
}

function idsEqual(
  left: string | null | undefined,
  right: string | null | undefined,
): boolean {
  const a = String(left ?? "").trim();
  const b = String(right ?? "").trim();
  return Boolean(a) && a === b;
}

/**
 * Native <select> placeholder. An empty-string option is treated as "already
 * selected" in some browsers, so picking the first real company never fires
 * onChange. A sentinel value makes that first pick a real value change.
 */
const SELECT_NONE = "__none__";

function isSelectNone(value: string | null | undefined): boolean {
  const text = String(value ?? "").trim();
  return !text || text === SELECT_NONE;
}

/** Drop duplicate SharePoint company rows so select keys/values stay unique. */
function uniqueCompanies(companies: Company[]): Company[] {
  const seenIds = new Set<string>();
  const seenNames = new Set<string>();
  const unique: Company[] = [];
  for (const company of companies) {
    const id = String(company.id ?? "").trim();
    const nameKey = company.companyName.trim().toLowerCase();
    if (id && seenIds.has(id)) continue;
    if (nameKey && seenNames.has(nameKey)) continue;
    if (id) seenIds.add(id);
    if (nameKey) seenNames.add(nameKey);
    unique.push(company);
  }
  return unique.sort((a, b) => a.companyName.localeCompare(b.companyName));
}

function findCompanyById(
  companies: Company[],
  id: string | null | undefined,
): Company | undefined {
  const key = String(id ?? "").trim();
  if (!key || isSelectNone(key)) return undefined;
  return companies.find((company) => idsEqual(company.id, key));
}

function resolveFormCompany(
  form: FormState,
  companies: Company[],
): { id: string; name: string } {
  const selectedId = String(form.companyId ?? "").trim();
  const selectedName = String(form.companyName ?? "").trim();
  const byId = selectedId
    ? companies.find((company) => idsEqual(company.id, selectedId))
    : undefined;
  if (byId) {
    return { id: byId.id, name: byId.companyName };
  }
  if (/^\d+$/.test(selectedName)) {
    const numeric = companies.find((company) => idsEqual(company.id, selectedName));
    if (numeric) {
      return { id: numeric.id, name: numeric.companyName };
    }
  }
  const byName = selectedName
    ? companies.find(
        (company) =>
          company.companyName.trim().toLowerCase() === selectedName.toLowerCase(),
      )
    : undefined;
  return {
    id: byName?.id ?? "",
    name: byName?.companyName ?? selectedName,
  };
}

function workforceMatchesCompany(
  row: AdminWorkforceOption,
  companyId: string,
  companyName: string,
): boolean {
  if (idsEqual(row.companyId, companyId)) return true;
  const name = companyName.trim().toLowerCase();
  return Boolean(name) && row.companyName.trim().toLowerCase() === name;
}

function applyCompanyLinkToForm(
  state: FormState,
  companies: Company[],
  row?: Record<string, unknown>,
): FormState {
  const lookupId = String(
    row?.companyLookupId ?? row?.companyId ?? state.companyId ?? "",
  ).trim();
  if (lookupId) {
    state.companyId = lookupId;
  }
  const resolved = resolveFormCompany(state, companies);
  if (resolved.id) {
    state.companyId = resolved.id;
  }
  if (resolved.name) {
    state.companyName = resolved.name;
  }
  if (resolved.id && !String(state.companyNumber ?? "").trim()) {
    const match = companies.find((company) => idsEqual(company.id, resolved.id));
    if (match?.companyNumber?.trim()) {
      state.companyNumber = match.companyNumber.trim();
    }
  }
  return state;
}

function matchWorkforceId(
  form: FormState,
  workforce: AdminWorkforceOption[],
  companies: Company[] = [],
): string {
  const name = String(form.candidateName ?? "").trim().toLowerCase();
  if (!name) return "";
  const company = resolveFormCompany(form, companies);
  const inCompany = workforce.filter((row) =>
    workforceMatchesCompany(row, company.id, company.name),
  );
  const pool = company.id || company.name ? inCompany : workforce;
  return (
    pool.find((row) => row.candidateName.trim().toLowerCase() === name)?.id ??
    (company.id || company.name
      ? ""
      : (workforce.find((row) => row.candidateName.trim().toLowerCase() === name)
          ?.id ?? ""))
  );
}

/** Semicolon-separated Workforce ids for multi-candidate create. */
function parseWorkforceIdList(raw: unknown): string[] {
  return String(raw ?? "")
    .split(/[;,|]+/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function formatWorkforceIdList(ids: string[]): string {
  return [...new Set(ids.map((id) => id.trim()).filter(Boolean))].join(";");
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
  companies = EMPTY_COMPANIES,
  workforce = EMPTY_WORKFORCE,
  permissionPeople = EMPTY_PERMISSION_PEOPLE,
  departments = EMPTY_DEPARTMENTS,
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
  warnings = EMPTY_WARNINGS,
  extraActions,
  breadcrumbs,
  wideTable = false,
  stickyLeadColumns = false,
  stickyColumnKey,
  tableClassName,
  getCreateDefaults,
  optimistic = false,
  optimisticRowLabel,
}: AdminCrudPageProps<T>) {
  const { pushToast } = useAdminToast();
  const router = useRouter();
  const pathname = usePathname();
  const searchParams = useSearchParams();
  const [rows, setRows] = useState<T[]>(initialRows);
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState("");
  const [companyFilter, setCompanyFilter] = useState(
    () => searchParams.get("company") ?? "",
  );
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
  /**
   * Quick-add for role-filtered dropdowns (Training Manager / Supervisor).
   * When non-null the QuickAddPermissionPersonModal is open; on save the
   * caller field auto-selects the new person and the permission people list
   * is refreshed. Null when no quick-add is in progress.
   */
  const [quickAdd, setQuickAdd] = useState<{
    fieldName: string;
    role: "Training Manager" | "Supervisor" | "Candidate";
    companyId: string;
    companyName: string;
  } | null>(null);
  /**
   * Optimistic UI: rows whose background API call failed after we already
   * updated the client. The pill above the table lets the admin see and
   * retry them without a red toast interrupting their flow.
   */
  const [reviewOps, setReviewOps] = useState<
    Array<{
      key: string;
      op: "delete" | "update" | "create";
      label: string;
      errorMessage: string;
      when: number;
      retry: () => void;
    }>
  >([]);
  const [reviewOpen, setReviewOpen] = useState(false);
  /**
   * Session-persisted set: IDs the user has already asked to delete but whose
   * SharePoint cascade hasn't confirmed yet. Refreshing the /admin/companies
   * page picks these up so the row does NOT come back mid-cascade. Auto-expires
   * after HIDDEN_TTL_MS to avoid hiding a row forever if the API silently died.
   */
  const HIDDEN_TTL_MS = 10 * 60 * 1000;
  const hiddenStorageKey = `admin-crud-hidden::${listUrl}`;
  const [hiddenIds, setHiddenIds] = useState<Set<string>>(() => new Set());
  const rowLabelFor = useCallback(
    (row: T): string => {
      if (optimisticRowLabel) return optimisticRowLabel(row);
      const anyRow = row as unknown as Record<string, unknown>;
      const candidates = [
        anyRow.companyName,
        anyRow.candidateName,
        anyRow.name,
        anyRow.title,
      ];
      for (const value of candidates) {
        if (typeof value === "string" && value.trim()) return value.trim();
      }
      return `#${row.id}`;
    },
    [optimisticRowLabel],
  );

  // Load persisted "hidden" IDs on first mount and prune expired ones.
  useEffect(() => {
    if (!optimistic) return;
    if (typeof window === "undefined") return;
    try {
      const raw = window.sessionStorage.getItem(hiddenStorageKey);
      if (!raw) return;
      const parsed = JSON.parse(raw) as Record<string, number>;
      const now = Date.now();
      const kept = new Set<string>();
      const pruned: Record<string, number> = {};
      for (const [id, when] of Object.entries(parsed)) {
        if (typeof when !== "number") continue;
        if (now - when > HIDDEN_TTL_MS) continue;
        kept.add(id);
        pruned[id] = when;
      }
      setHiddenIds(kept);
      window.sessionStorage.setItem(hiddenStorageKey, JSON.stringify(pruned));
    } catch {
      // sessionStorage disabled / quota — safe to fall back to in-memory only.
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hiddenStorageKey, optimistic]);

  const persistHidden = useCallback(
    (next: Set<string>) => {
      if (typeof window === "undefined") return;
      try {
        const map: Record<string, number> = {};
        const now = Date.now();
        for (const id of next) map[id] = now;
        window.sessionStorage.setItem(hiddenStorageKey, JSON.stringify(map));
      } catch {
        // ignore
      }
    },
    [hiddenStorageKey],
  );

  const addHiddenId = useCallback(
    (id: string) => {
      setHiddenIds((current) => {
        if (current.has(id)) return current;
        const next = new Set(current);
        next.add(id);
        persistHidden(next);
        return next;
      });
    },
    [persistHidden],
  );

  const removeHiddenId = useCallback(
    (id: string) => {
      setHiddenIds((current) => {
        if (!current.has(id)) return current;
        const next = new Set(current);
        next.delete(id);
        persistHidden(next);
        return next;
      });
    },
    [persistHidden],
  );

  // Apply session-persisted hidden IDs to the SSR payload on mount so a
  // hard refresh mid-cascade doesn't briefly re-show a row the user has
  // already asked to delete.
  useEffect(() => {
    if (!optimistic) return;
    if (hiddenIds.size === 0) return;
    setRows((current) => {
      const filtered = current.filter((row) => !hiddenIds.has(row.id));
      if (filtered.length === current.length) return current;
      onRowsChange?.(filtered);
      return filtered;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hiddenIds, optimistic]);
  // Prefer live API refresh when opening create/edit; fall back to prop.
  // Do not sync prop→state in an effect — unstable array identities (or
  // default `= []`) re-trigger setState every render and hit max update depth.
  const [permissionPeopleLive, setPermissionPeopleLive] = useState<
    AdminPermissionPersonOption[] | null
  >(null);
  const livePermissionPeople = permissionPeopleLive ?? permissionPeople;
  const needsPermissionPeople = fields.some(
    (field) =>
      Boolean(field.permissionRoleFilter) ||
      Boolean(field.sharePointRoleTypeFilter),
  );
  const openedFromQueryRef = useRef(false);

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
      setPermissionPeopleLive(
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
    const next = buildInitialForm(fields);
    const defaults = getCreateDefaults?.(rows) ?? {};
    for (const [key, value] of Object.entries(defaults)) {
      if (value !== undefined) next[key] = value;
    }
    if (enableCompanyFilter && companyFilter.trim()) {
      next.companyName = companyFilter.trim();
    }
    applyCompanyLinkToForm(next, companies);
    setForm(next);
    setFormError(null);
    setWorkforceQuery("");
    void refreshPermissionPeople();
    setDrawerOpen(true);
  }

  const createAction = searchParams.get("action");
  useEffect(() => {
    if (openedFromQueryRef.current) return;
    if (
      !allowCreate ||
      (createAction !== "add" &&
        createAction !== "create" &&
        createAction !== "upload")
    ) {
      return;
    }
    openedFromQueryRef.current = true;
    openCreate();
    router.replace(pathname);
    // openCreate reads latest form helpers; guard ref prevents re-entry.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allowCreate, createAction, pathname, router]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const response = await fetch(listUrl, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(await readError(response));
      }
      const payload = await response.json();
      let nextRows = mapResponse(payload);
      // Optimistic: if a row is still "hidden" (its background delete has
      // not confirmed on SharePoint yet), keep it out of the visible list.
      // Once SharePoint returns without the row, we drop it from hiddenIds.
      if (optimistic && hiddenIds.size > 0) {
        const seenInPayload = new Set(nextRows.map((row) => row.id));
        nextRows = nextRows.filter((row) => !hiddenIds.has(row.id));
        for (const id of hiddenIds) {
          if (!seenInPayload.has(id)) removeHiddenId(id);
        }
      }
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
  }, [
    listUrl,
    mapResponse,
    onRowsChange,
    pushToast,
    optimistic,
    hiddenIds,
    removeHiddenId,
  ]);

  const patchRow = useCallback(
    (id: string, patch: Partial<T>) => {
      setRows((current) => {
        const next = current.map((row) =>
          row.id === id ? { ...row, ...patch } : row,
        );
        onRowsChange?.(next);
        return next;
      });
    },
    [onRowsChange],
  );

  const columnRenderCtx = useMemo(
    () => ({ reload: load, patchRow }),
    [load, patchRow],
  );

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

  /**
   * Runs the DELETE fetch. Shared between the synchronous path and the
   * optimistic (background) path.
   */
  async function performDeleteRequest(id: string) {
    if (!deleteUrl) return;
    const response = await fetch(deleteUrl(id), { method: "DELETE" });
    if (!response.ok) throw new Error(await readError(response));
    await response.json().catch(() => null);
  }

  async function deleteOne(id: string) {
    if (!deleteUrl) return;
    const ok = window.confirm(
      `Delete this record?${deleteConfirmExtra ? `\n\n${deleteConfirmExtra}` : "\n\nThis cannot be undone."}`,
    );
    if (!ok) return;

    // Optimistic path: hide the row immediately, run DELETE silently in the
    // background. On failure we silently put the row back and log it to the
    // review pill above the table.
    if (optimistic) {
      const snapshot = rows.find((row) => row.id === id);
      if (!snapshot) return;
      setSelectedIds((current) => {
        const next = new Set(current);
        next.delete(id);
        return next;
      });
      setRows((current) => {
        const next = current.filter((row) => row.id !== id);
        onRowsChange?.(next);
        return next;
      });
      // Keep this ID hidden across a page refresh until SharePoint confirms.
      addHiddenId(id);
      if (editing?.id === id) {
        setDrawerOpen(false);
        setEditing(null);
        openedFromQueryRef.current = false;
      }
      // Immediate confirmation for the user — same tone as the sync path.
      pushToast("Record deleted", "success");
      const restore = () => {
        setRows((current) => {
          if (current.some((row) => row.id === id)) return current;
          const next = [...current, snapshot];
          onRowsChange?.(next);
          return next;
        });
      };
      void (async () => {
        try {
          await performDeleteRequest(id);
          // Reconcile with SharePoint truth after the cascade settles.
          void load();
        } catch (error) {
          restore();
          removeHiddenId(id);
          const message =
            error instanceof Error ? error.message : "Delete failed";
          const reviewKey = `delete-${id}-${Date.now()}`;
          setReviewOps((prev) => [
            ...prev,
            {
              key: reviewKey,
              op: "delete",
              label: rowLabelFor(snapshot),
              errorMessage: message,
              when: Date.now(),
              retry: () => {
                setReviewOps((current) =>
                  current.filter((entry) => entry.key !== reviewKey),
                );
                void deleteOne(id);
              },
            },
          ]);
        }
        // On success, keep the hidden-ID entry a little longer so a fast
        // reload right after the cascade completes doesn't briefly show the
        // row before SharePoint's list read reflects the delete. It will
        // naturally age out via HIDDEN_TTL_MS pruning on next mount.
      })();
      return;
    }

    setDeleting(true);
    try {
      const response = await fetch(deleteUrl(id), { method: "DELETE" });
      if (!response.ok) throw new Error(await readError(response));
      await response.json().catch(() => null);
      pushToast("Record deleted", "success");
      setSelectedIds((current) => {
        const next = new Set(current);
        next.delete(id);
        return next;
      });
      // Drop locally first so a stale list payload cannot put the row back.
      setRows((current) => {
        const next = current.filter((row) => row.id !== id);
        onRowsChange?.(next);
        return next;
      });
      if (editing?.id === id) {
        setDrawerOpen(false);
        setEditing(null);
        openedFromQueryRef.current = false;
      }
      // Reload, then force-drop the id again (guards against cache races).
      await load();
      setRows((current) => {
        const next = current.filter((row) => row.id !== id);
        onRowsChange?.(next);
        return next;
      });
      router.refresh();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Delete failed",
        "error",
      );
      // Put fresh SharePoint truth back if delete failed mid-flight.
      await load();
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

    // Optimistic bulk delete: hide every selected row instantly, keep them
    // hidden across page refresh (sessionStorage), show the pending pill.
    // POST /bulk-delete runs in the background; on failure any rows that
    // couldn't be deleted come silently back with a review entry.
    if (optimistic) {
      const snapshotById = new Map<string, T>();
      for (const id of ids) {
        const row = rows.find((r) => r.id === id);
        if (row) snapshotById.set(id, row);
      }
      const targetIds = Array.from(snapshotById.keys());
      if (targetIds.length === 0) return;
      setSelectedIds(new Set());
      setRows((current) => {
        const target = new Set(targetIds);
        const next = current.filter((row) => !target.has(row.id));
        onRowsChange?.(next);
        return next;
      });
      setHiddenIds((current) => {
        const next = new Set(current);
        for (const id of targetIds) next.add(id);
        persistHidden(next);
        return next;
      });
      pushToast(
        `Deleted ${targetIds.length} record${targetIds.length === 1 ? "" : "s"}`,
        "success",
      );

      void (async () => {
        try {
          const response = await fetch(bulkDeleteUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ ids: targetIds }),
          });
          if (!response.ok) throw new Error(await readError(response));
          void load();
        } catch (error) {
          // Full-batch failure — restore everything silently. Partial failures
          // (some ids deleted, some not) are handled by the reconciling load()
          // call above: it sees which rows SharePoint no longer has and
          // clears their hidden entries; the rest naturally come back.
          const message =
            error instanceof Error ? error.message : "Bulk delete failed";
          const reviewKey = `bulk-delete-${Date.now()}`;
          const failedIds = Array.from(snapshotById.keys());
          setRows((current) => {
            const known = new Set(current.map((row) => row.id));
            const restored = [...current];
            for (const id of failedIds) {
              if (known.has(id)) continue;
              const row = snapshotById.get(id);
              if (row) restored.push(row);
            }
            onRowsChange?.(restored);
            return restored;
          });
          setHiddenIds((current) => {
            const next = new Set(current);
            for (const id of failedIds) next.delete(id);
            persistHidden(next);
            return next;
          });
          setReviewOps((prev) => [
            ...prev,
            {
              key: reviewKey,
              op: "delete",
              label: `${failedIds.length} record${failedIds.length === 1 ? "" : "s"}`,
              errorMessage: message,
              when: Date.now(),
              retry: () => {
                setReviewOps((current) =>
                  current.filter((entry) => entry.key !== reviewKey),
                );
                setSelectedIds(new Set(failedIds));
                void deleteSelected();
              },
            },
          ]);
        }
      })();
      return;
    }

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
    const next = withAutoFilledBlankExpiry(
      fields,
      applyCompanyLinkToForm(
        buildInitialForm(fields, row as unknown as Record<string, unknown>),
        companies,
        row as unknown as Record<string, unknown>,
      ),
    );
    setForm(next);
    setFormError(null);
    setWorkforceQuery("");
    void refreshPermissionPeople();
    setDrawerOpen(true);
  }

  function validate(next: FormState): string | null {
    for (const field of fields) {
      const value = next[field.name];
      if (
        field.type === "workforce" &&
        field.workforceMultiSelect &&
        !editing
      ) {
        if (field.required && parseWorkforceIdList(next.candidateWorkforceIds).length === 0) {
          return "Select at least one candidate.";
        }
        continue;
      }
      if (field.required && field.type !== "boolean") {
        if (
          value === "" ||
          value === null ||
          value === undefined ||
          isSelectNone(String(value))
        ) {
          return `${field.label} is required.`;
        }
      }
      if (
        field.type === "email" &&
        typeof value === "string" &&
        value.trim() &&
        !isValidEmail(value)
      ) {
        return `${field.label} must be a valid email address (e.g. name@company.org).`;
      }
    }
    return null;
  }

  async function persist(next: FormState) {
    setSaving(true);
    setFormError(null);
    try {
      const body: Record<string, unknown> = { ...next };
      if (typeof body.permissionRole === "string") {
        body.permissionRole = body.permissionRole.trim();
      }
      if (
        body.permissionRole === "Customer" ||
        body.permissionRole === "Candidate"
      ) {
        body.accessScope = "Candidate Only";
      }
      // Routing bucket RoleType is Admin | Customer. Never send it as the
      // form role — it made every Supervisor save as Customer.
      delete body.roleType;
      if (
        shouldAutoFillBlankExpiry(fields, String(body.trainingOutcome ?? "")) &&
        !String(body.expiry ?? "").trim()
      ) {
        const filled = autoExpiryIso(fields, String(body.trainingDate ?? ""));
        if (filled) body.expiry = filled;
      }
      if (typeof body.companyId === "string" && isSelectNone(body.companyId)) {
        body.companyId = "";
      }
      if (
        !String(body.accessScope ?? "")
          .toLowerCase()
          .includes("department")
      ) {
        body.departmentsAllowed = "";
      }

      // Cross-check: if the form declares a companyId, force companyName to
      // match the current companies list. This defends against two failure
      // modes at once: (a) a stale companyName lingering from a previous
      // edit session while a fresh companyId was picked, and (b) a mismatch
      // between the displayed dropdown label and the underlying id (which
      // could happen if the browser dropdown ever failed to register a click
      // and left the form partially updated).
      const linkedCompany = resolveFormCompany(next, companies);
      if (linkedCompany.id) {
        body.companyId = linkedCompany.id;
        body.companyName = linkedCompany.name;
      } else if (body.companyId && companies.length) {
        const company = companies.find((item) =>
          idsEqual(item.id, String(body.companyId)),
        );
        if (!company) {
          throw new Error(
            "Selected company is no longer available. Refresh the page and pick a company again.",
          );
        }
        body.companyName = company.companyName;
      } else if (
        fields.some(
          (field) => field.type === "company" && field.name === "companyId" && field.required,
        ) &&
        !body.companyId
      ) {
        // Required-company field somehow reached persist without a value — hard
        // stop so we never write a permission row that could be assigned to the
        // wrong company via a fallback lookup on companyName alone.
        throw new Error(
          "Please pick a company from the Company dropdown before saving.",
        );
      }
      const workforceId = matchWorkforceId(
        {
          ...next,
          companyId: String(body.companyId ?? next.companyId ?? ""),
          companyName: String(body.companyName ?? next.companyName ?? ""),
        },
        workforce,
        companies,
      );
      if (workforceId) {
        body.workforceId = workforceId;
      }

      const isCreate = !editing;
      if (!isCreate && !updateUrl) {
        throw new Error("Editing is not available for this list.");
      }

      const multiWorkforceField = fields.find(
        (field) => field.type === "workforce" && field.workforceMultiSelect,
      );
      const multiIds =
        isCreate && multiWorkforceField
          ? parseWorkforceIdList(next.candidateWorkforceIds)
          : [];

      if (multiIds.length > 1) {
        delete body.candidateWorkforceIds;
        const created: string[] = [];
        const failures: string[] = [];
        for (const id of multiIds) {
          const hit = workforce.find((row) => row.id === id);
          if (!hit) {
            failures.push(`Unknown Workforce id ${id}`);
            continue;
          }
          const perBody: Record<string, unknown> = {
            ...body,
            workforceId: hit.id,
            candidateName: hit.candidateName,
            companyName: hit.companyName || body.companyName,
            companyId: hit.companyId || body.companyId,
            nporsNumber: hit.nporsNumbers || "",
            eusrNumber: hit.eusrNumber || "",
            swqrNumber: hit.swqrNumber || "",
            inHouseCertificationNumber: hit.inHouseCertificationNumber || "",
            workforceNumber: hit.workforceNumber || "",
            niNumber: hit.niNumber || "",
          };
          try {
            const response = await fetch(createUrl ?? listUrl, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(perBody),
            });
            if (!response.ok) {
              failures.push(
                `${hit.candidateName}: ${await readError(response)}`,
              );
              continue;
            }
            created.push(hit.candidateName);
          } catch (error) {
            failures.push(
              `${hit.candidateName}: ${
                error instanceof Error ? error.message : "Save failed."
              }`,
            );
          }
        }
        if (created.length === 0) {
          throw new Error(
            failures[0] || "Could not save training for any selected candidate.",
          );
        }
        const syncNote =
          failures.length > 0
            ? ` ${failures.length} failed: ${failures.slice(0, 3).join("; ")}${failures.length > 3 ? "…" : ""}`
            : "";
        pushToast(
          `Created ${created.length} NPORS record${created.length === 1 ? "" : "s"}.${syncNote}`,
          failures.length > 0 ? "error" : "success",
        );
        if (failures.length > 0) {
          setFormError(
            `Saved ${created.length}, but ${failures.length} failed. ${failures.slice(0, 5).join(" · ")}`,
          );
        } else {
          setDrawerOpen(false);
          openedFromQueryRef.current = false;
        }
        await load();
        router.refresh();
        return;
      }

      if (multiIds.length === 1) {
        const hit = workforce.find((row) => row.id === multiIds[0]);
        if (hit) {
          body.workforceId = hit.id;
          body.candidateName = hit.candidateName;
          body.nporsNumber = hit.nporsNumbers || body.nporsNumber || "";
          body.eusrNumber = hit.eusrNumber || body.eusrNumber || "";
          body.swqrNumber = hit.swqrNumber || body.swqrNumber || "";
          body.workforceNumber =
            hit.workforceNumber || body.workforceNumber || "";
          body.niNumber = hit.niNumber || body.niNumber || "";
        }
      }
      delete body.candidateWorkforceIds;

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
            warnings?: number;
          };
          items?: Array<{
            warnings?: string[];
            errors?: string[];
            skipped?: boolean;
            skipReason?: string;
            fieldsUpdated?: string[];
          }>;
        };
      } | null;
      const sync = payload?.matrixSync?.summary;
      const syncItem = payload?.matrixSync?.items?.[0];
      const matrixTouched =
        (sync?.updated ?? 0) + (sync?.created ?? 0) > 0 ||
        (syncItem?.fieldsUpdated?.length ?? 0) > 0;
      const syncNote = sync
        ? matrixTouched
          ? ` Matrix sync: ${sync.updated ?? 0} updated, ${sync.created ?? 0} created.`
          : ` Matrix already up to date.`
        : "";
      pushToast(
        (isCreate ? "Record created." : "Record updated.") + syncNote,
        "success",
      );
      setDrawerOpen(false);
      // Release the one-shot ?action=add guard so subsequent
      // "+ Add …" URL-driven opens still work in the same session.
      openedFromQueryRef.current = false;
      await load();
      router.refresh();
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

  function commitSelectValue(field: AdminFieldConfig, rawValue: string) {
    const value = isSelectNone(rawValue) ? "" : rawValue;
    if (
      field.type === "company" &&
      (field.name === "companyName" || field.name === "companyId")
    ) {
      setWorkforceQuery("");
      void refreshPermissionPeople();
    }
    setForm((current) => {
      if (field.type !== "company") {
        const next: FormState = {
          ...current,
          [field.name]: value,
        };
        if (
          field.name === "trainingOutcome" &&
          shouldAutoFillBlankExpiry(fields, value) &&
          !String(current.expiry ?? "").trim()
        ) {
          const filled = autoExpiryIso(
            fields,
            String(current.trainingDate ?? ""),
          );
          if (filled) next.expiry = filled;
        }
        if (
          field.name === "permissionRole" &&
          (value === "Customer" || value === "Candidate")
        ) {
          next.accessScope = "Candidate Only";
        }
        return next;
      }

      const next: FormState = {
        ...current,
        [field.name]: value,
      };

      if (field.name === "companyName" || field.name === "companyId") {
        next.trainingManager = "";
        next.supervisor = "";
        next.department = "";
        next.departmentsAllowed = "";
        next.candidateName = "";
        next.candidateWorkforceIds = "";
        next.nporsNumber = "";
        next.eusrNumber = "";
        next.swqrNumber = "";
        next.workforceNumber = "";
        next.niNumber = "";
        next.inHouseCertificationNumber = "";
      }

      if (field.name === "companyName") {
        const companyMatch = companies.find(
          (company) =>
            company.companyName.trim().toLowerCase() ===
            value.trim().toLowerCase(),
        );
        next.companyId = companyMatch?.id ?? "";
        next.companyNumber = companyMatch?.companyNumber?.trim() || "";
      }

      if (field.name === "companyId") {
        const companyMatch = findCompanyById(companies, value);
        next.companyNumber = companyMatch?.companyNumber?.trim() || "";
        next.companyName = companyMatch?.companyName ?? "";
      }

      if (field.name !== "companyName") {
        return next;
      }

      const candidateName = String(current.candidateName ?? "")
        .trim()
        .toLowerCase();
      const currentCompany = String(current.companyName ?? "")
        .trim()
        .toLowerCase();
      const selectedCandidate =
        workforce.find(
          (row) =>
            row.candidateName.trim().toLowerCase() === candidateName &&
            row.companyName.trim().toLowerCase() === currentCompany,
        ) ??
        workforce.find(
          (row) => row.candidateName.trim().toLowerCase() === candidateName,
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
        next.niNumber = "";
      }

      return next;
    });
  }

  function requestSave() {
    const next: FormState = { ...form };
    if (
      fields.some((field) => field.name === "expiry") &&
      !String(next.expiry ?? "").trim() &&
      shouldAutoFillBlankExpiry(
        fields,
        String(next.trainingOutcome ?? ""),
      )
    ) {
      const filled = autoExpiryIso(
        fields,
        String(next.trainingDate ?? ""),
      );
      if (filled) next.expiry = filled;
      setForm(next);
    }
    const error = validate(next);
    if (error) {
      setFormError(error);
      return;
    }

    if (
      confirmInactive &&
      typeof next.status === "string" &&
      next.status.toLowerCase() === "inactive"
    ) {
      setPendingSave(next);
      setConfirmOpen(true);
      return;
    }

    void persist(next);
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

      {reviewOps.length > 0 ? (
        <div className={styles.reviewBar}>
          <span aria-hidden="true">⚠</span>
          <span>
            {reviewOps.length} action
            {reviewOps.length === 1 ? "" : "s"} need review
          </span>
          <button
            type="button"
            className={styles.reviewBarButton}
            onClick={() => setReviewOpen((value) => !value)}
          >
            {reviewOpen ? "Hide" : "Review"}
          </button>
          <button
            type="button"
            className={styles.reviewBarClear}
            onClick={() => {
              setReviewOps([]);
              setReviewOpen(false);
            }}
          >
            Dismiss
          </button>
        </div>
      ) : null}
      {reviewOpen && reviewOps.length > 0 ? (
        <div className={styles.reviewPanel} role="dialog" aria-label="Pending actions">
          <ul className={styles.reviewList}>
            {reviewOps.map((entry) => (
              <li key={entry.key} className={styles.reviewItem}>
                <div>
                  <p className={styles.reviewItemTitle}>
                    {entry.op === "delete"
                      ? "Delete"
                      : entry.op === "update"
                        ? "Update"
                        : "Create"}
                    : {entry.label}
                  </p>
                  <p className={styles.reviewItemError}>{entry.errorMessage}</p>
                </div>
                <div className={styles.reviewItemActions}>
                  <button
                    type="button"
                    className={styles.reviewItemRetry}
                    onClick={entry.retry}
                  >
                    Retry
                  </button>
                  <button
                    type="button"
                    className={styles.reviewItemDismiss}
                    onClick={() =>
                      setReviewOps((current) =>
                        current.filter((row) => row.key !== entry.key),
                      )
                    }
                  >
                    Dismiss
                  </button>
                </div>
              </li>
            ))}
          </ul>
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
              {uniqueCompanies(companies).map((company) => (
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
                <th
                  scope="col"
                  className={
                    updateUrl || deleteUrl ? styles.stickyActionsCell : undefined
                  }
                >
                  Actions
                </th>
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
                      {column.render(row, columnRenderCtx)}
                    </td>
                  ))}
                  <td
                    className={
                      updateUrl || deleteUrl
                        ? styles.stickyActionsCell
                        : undefined
                    }
                  >
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
        onClose={() => {
          setDrawerOpen(false);
          // Release the one-shot ?action=add URL guard so the next
          // "+ Add Training Manager" / "+ Add Customer" button click re-opens
          // the drawer. Without this reset the guard stayed true for the
          // whole page lifetime and subsequent button clicks silently no-op.
          openedFromQueryRef.current = false;
        }}
        wide={drawerWide}
        footer={
          <>
            {editing && deleteUrl ? (
              <button
                type="button"
                className={styles.linkButtonDanger}
                disabled={saving || deleting}
                onClick={() => {
                  void deleteOne(editing.id);
                }}
              >
                {deleting ? "Deleting…" : "Delete"}
              </button>
            ) : null}
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => {
                setDrawerOpen(false);
                openedFromQueryRef.current = false;
              }}
              disabled={saving || deleting}
            >
              Cancel
            </button>
            <button
              type="button"
              className={styles.primaryButton}
              disabled={saving || deleting}
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
              const selectedId = matchWorkforceId(form, workforce, companies);
              const resolvedCompany = resolveFormCompany(form, companies);
              const selectedCompanyName = resolvedCompany.name.trim().toLowerCase();
              const resolvedCompanyId = resolvedCompany.id;
              const hasCompany = Boolean(
                selectedCompanyName || resolvedCompanyId,
              );
              const query = workforceQuery.trim().toLowerCase();
              const companyWorkforce = hasCompany
                ? workforce.filter((row) =>
                    workforceMatchesCompany(
                      row,
                      resolvedCompanyId,
                      resolvedCompany.name,
                    ),
                  )
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
                      row.niNumber,
                    ]
                      .filter(Boolean)
                      .join(" ")
                      .toLowerCase();
                    return haystack.includes(query);
                  })
                : companyWorkforce;
              const candidateDisabled = field.readOnly || !hasCompany;
              const useMulti =
                Boolean(field.workforceMultiSelect) && !editing;
              const selectedMultiIds = new Set(
                parseWorkforceIdList(form.candidateWorkforceIds),
              );
              const emptyOptionLabel = !hasCompany
                ? "Select a company first…"
                : companyWorkforce.length === 0
                  ? "No Workforce candidates for this company"
                  : query && filteredWorkforce.length === 0
                    ? `No matches for “${workforceQuery.trim()}”`
                    : "Select candidate…";

              function applyWorkforceHit(hit: AdminWorkforceOption | null) {
                if (!hit) {
                  setForm((current) => ({
                    ...current,
                    candidateName: "",
                    candidateWorkforceIds: "",
                    companyName: current.companyName,
                    companyNumber: current.companyNumber,
                    nporsNumber: "",
                    eusrNumber: "",
                    swqrNumber: "",
                    inHouseCertificationNumber: "",
                    workforceNumber: "",
                    niNumber: "",
                  }));
                  return;
                }
                const companyMatch =
                  companies.find((company) =>
                    idsEqual(company.id, hit.companyId),
                  ) ??
                  companies.find(
                    (company) =>
                      company.companyName.trim().toLowerCase() ===
                      hit.companyName.trim().toLowerCase(),
                  );
                setForm((current) => ({
                  ...current,
                  candidateName: hit.candidateName,
                  candidateWorkforceIds: hit.id,
                  companyName: hit.companyName,
                  companyId: hit.companyId || current.companyId,
                  companyNumber:
                    companyMatch?.companyNumber?.trim() ||
                    String(current.companyNumber ?? ""),
                  nporsNumber: hit.nporsNumbers || "",
                  eusrNumber: hit.eusrNumber || "",
                  swqrNumber: hit.swqrNumber || "",
                  inHouseCertificationNumber:
                    hit.inHouseCertificationNumber || "",
                  workforceNumber: hit.workforceNumber || "",
                  niNumber: hit.niNumber || "",
                }));
              }

              function toggleMultiWorkforce(id: string, checked: boolean) {
                setForm((current) => {
                  const prev = parseWorkforceIdList(
                    current.candidateWorkforceIds,
                  );
                  const nextIds = checked
                    ? formatWorkforceIdList([...prev, id])
                    : formatWorkforceIdList(prev.filter((entry) => entry !== id));
                  const ordered = parseWorkforceIdList(nextIds)
                    .map((entry) => workforce.find((row) => row.id === entry))
                    .filter((row): row is AdminWorkforceOption => Boolean(row));
                  const first = ordered[0] ?? null;
                  return {
                    ...current,
                    candidateWorkforceIds: nextIds,
                    candidateName: ordered
                      .map((row) => row.candidateName)
                      .join("; "),
                    nporsNumber:
                      ordered.length === 1 ? first?.nporsNumbers || "" : "",
                    eusrNumber:
                      ordered.length === 1 ? first?.eusrNumber || "" : "",
                    swqrNumber:
                      ordered.length === 1 ? first?.swqrNumber || "" : "",
                    inHouseCertificationNumber:
                      ordered.length === 1
                        ? first?.inHouseCertificationNumber || ""
                        : "",
                    workforceNumber:
                      ordered.length === 1
                        ? first?.workforceNumber || ""
                        : ordered.length > 1
                          ? `${ordered.length} selected`
                          : "",
                    niNumber: ordered.length === 1 ? first?.niNumber || "" : "",
                  };
                });
              }

              control = (
                <fieldset className={styles.field}>
                  <legend className={styles.fieldLabel}>
                    {useMulti
                      ? `${field.label}s (same company)`
                      : field.label}
                    {hasCompany
                      ? useMulti
                        ? " — tick everyone who took this course"
                        : " (scoped to selected company)"
                      : " — select a company first"}
                  </legend>
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
                  {useMulti ? (
                    <>
                      {hasCompany && filteredWorkforce.length > 0 ? (
                        <div className={styles.workforceMultiActions}>
                          <button
                            type="button"
                            className={styles.linkButton}
                            disabled={candidateDisabled}
                            onClick={() => {
                              const ids = filteredWorkforce.map((row) => row.id);
                              setForm((current) => {
                                const merged = formatWorkforceIdList([
                                  ...parseWorkforceIdList(
                                    current.candidateWorkforceIds,
                                  ),
                                  ...ids,
                                ]);
                                const ordered = parseWorkforceIdList(merged)
                                  .map((entry) =>
                                    workforce.find((row) => row.id === entry),
                                  )
                                  .filter(
                                    (row): row is AdminWorkforceOption =>
                                      Boolean(row),
                                  );
                                return {
                                  ...current,
                                  candidateWorkforceIds: merged,
                                  candidateName: ordered
                                    .map((row) => row.candidateName)
                                    .join("; "),
                                  nporsNumber: "",
                                  eusrNumber: "",
                                  swqrNumber: "",
                                  workforceNumber:
                                    ordered.length > 0
                                      ? `${ordered.length} selected`
                                      : "",
                                  niNumber: "",
                                  inHouseCertificationNumber: "",
                                };
                              });
                            }}
                          >
                            Select all shown
                          </button>
                          <button
                            type="button"
                            className={styles.linkButton}
                            disabled={
                              candidateDisabled || selectedMultiIds.size === 0
                            }
                            onClick={() => applyWorkforceHit(null)}
                          >
                            Clear
                          </button>
                          {selectedMultiIds.size > 0 ? (
                            <span className={styles.helpText}>
                              {selectedMultiIds.size} selected
                            </span>
                          ) : null}
                        </div>
                      ) : null}
                      <div className={styles.multiSelectList}>
                        {!hasCompany ? (
                          <p className={styles.helpText}>{emptyOptionLabel}</p>
                        ) : filteredWorkforce.length === 0 ? (
                          <p className={styles.helpText}>{emptyOptionLabel}</p>
                        ) : (
                          filteredWorkforce.map((row) => {
                            const checked = selectedMultiIds.has(row.id);
                            const label = row.workforceNumber?.trim()
                              ? `${row.candidateName} (${row.workforceNumber})`
                              : row.candidateName;
                            return (
                              <label
                                key={row.id}
                                className={styles.multiSelectOption}
                              >
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  disabled={candidateDisabled}
                                  onChange={(event) =>
                                    toggleMultiWorkforce(
                                      row.id,
                                      event.target.checked,
                                    )
                                  }
                                />
                                <span>{label}</span>
                              </label>
                            );
                          })
                        )}
                      </div>
                    </>
                  ) : (
                    <select
                      className={styles.select}
                      value={selectedId}
                      disabled={candidateDisabled}
                      onChange={(event) => {
                        const hit =
                          workforce.find(
                            (row) => row.id === event.target.value,
                          ) ?? null;
                        applyWorkforceHit(hit);
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
                  )}
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
                </fieldset>
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
              const scopeMode = field.companyScopedDepartments;
              let multiOptions = field.options ?? [];
              if (scopeMode === "all") {
                // Cross-company mode: every active department, tagged with its
                // company so the admin can tell duplicates apart.
                multiOptions = departments.map((dept) => ({
                  value: valueMode === "name" ? dept.name : dept.id,
                  label: dept.companyName
                    ? `${dept.name} — ${dept.companyName}`
                    : dept.name,
                }));
              } else if (scopeMode) {
                if (!selectedCompanyName && !selectedCompanyId) {
                  multiOptions = [];
                } else {
                  multiOptions = departments
                    .filter((dept) => {
                      if (
                        selectedCompanyId &&
                        idsEqual(dept.companyId, selectedCompanyId)
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
              if (scopeMode && editing) {
                // Keep already-selected values visible on edit even when the
                // caller only passed Active departments. Do not do this on
                // create — leftover ids from another company (e.g. CVA while
                // DBS is selected) would reappear and fail save.
                for (const value of selected) {
                  const key = value.trim().toLowerCase();
                  if (
                    !multiOptions.some(
                      (option) =>
                        option.value.trim().toLowerCase() === key ||
                        option.label.trim().toLowerCase() === key,
                    )
                  ) {
                    multiOptions = [
                      ...multiOptions,
                      { value, label: `${value} (inactive)` },
                    ];
                  }
                }
              }
              const hideDeptCoverage =
                field.name === "departmentsAllowed" &&
                !String(form.accessScope ?? "")
                  .toLowerCase()
                  .includes("department");
              control = (
                <fieldset className={styles.field}>
                  <legend className={styles.fieldLabel}>{field.label}</legend>
                  {hideDeptCoverage ? (
                    <p className={styles.helpText}>
                      Only used when Access scope is Department Only. Full
                      Company covers every department in the selected company.
                    </p>
                  ) : field.companyScopedDepartments === true &&
                    !selectedCompanyName &&
                    !selectedCompanyId ? (
                    <p className={styles.helpText}>
                      Select a company first to choose departments.
                    </p>
                  ) : null}
                  {!hideDeptCoverage &&
                  field.companyScopedDepartments === true &&
                  (selectedCompanyName || selectedCompanyId) &&
                  multiOptions.length === 0 ? (
                    <p className={styles.helpText}>
                      No departments for this company yet. Add them under
                      Departments.
                    </p>
                  ) : null}
                  {!hideDeptCoverage &&
                  field.companyScopedDepartments === "all" &&
                  multiOptions.length === 0 ? (
                    <p className={styles.helpText}>
                      No active departments exist yet. Add them under
                      Departments.
                    </p>
                  ) : null}
                  {hideDeptCoverage ? null : (
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
                  )}
                </fieldset>
              );
            } else if (
              field.name === "accessScope" &&
              (String(form.permissionRole ?? "") === "Customer" ||
                String(form.permissionRole ?? "") === "Candidate")
            ) {
              control = (
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>{field.label}</span>
                  <p className={styles.helpText}>
                    Customer accounts see their own profile only (matched to
                    Workforce by email).
                  </p>
                </label>
              );
            } else if (field.type === "company" || field.type === "select") {
              const selectedCompanyName = String(form.companyName ?? "")
                .trim()
                .toLowerCase();
              const selectedCompanyId = String(form.companyId ?? "").trim();

              let options: Array<{ value: string; label: string }>;
              if (field.type === "company") {
                options = uniqueCompanies(companies).map((company) => ({
                  value:
                    field.name === "companyId"
                      ? String(company.id)
                      : company.companyName,
                  label: company.companyName,
                }));
              } else if (field.sharePointRoleTypeFilter) {
                // STRICT: match the SharePoint RoleType exactly. Preferred
                // over `permissionRoleFilter` for pages that need to pick a
                // specific role (Training Manager, Supervisor, etc.) — the
                // routing-role bucket used by `permissionRoleFilter` conflates
                // Admin + Training Manager and Supervisor + Candidate.
                const target = field.sharePointRoleTypeFilter;
                const emptyLabel =
                  target === "Training Manager"
                    ? "No active Training Managers in Permissions for this company"
                    : target === "Supervisor"
                      ? "No active Supervisors in Permissions for this company"
                      : `No active ${target}s in Permissions for this company`;
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
                  const targetKey = target.toLowerCase().replace(/\s+/g, " ");
                  const scoped = livePermissionPeople.filter((person) => {
                    if ((person.status || "").toLowerCase() !== "active") {
                      return false;
                    }
                    const spRole = (person.sharePointRoleType || "")
                      .trim()
                      .toLowerCase()
                      .replace(/\s+/g, " ");
                    // Accept the exact SharePoint RoleType. Legacy rows with
                    // an empty RoleType are excluded from strict filters —
                    // they should be reclassified before appearing in a
                    // role-specific picker.
                    if (spRole !== targetKey) {
                      // Allow "TrainingManager" as a legacy spelling.
                      if (
                        !(target === "Training Manager" && spRole === "trainingmanager")
                      ) {
                        return false;
                      }
                    }
                    if (
                      resolvedCompanyId &&
                      person.companyId &&
                      person.companyId !== resolvedCompanyId
                    ) {
                      return false;
                    }
                    if (
                      !person.companyId &&
                      selectedCompanyName &&
                      (person.companyName || "").trim().toLowerCase() !==
                        selectedCompanyName
                    ) {
                      return false;
                    }
                    return true;
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
                  options = withAssignedWorkforcePersonOption(
                    peopleOptions,
                    String(form[field.name] ?? ""),
                    livePermissionPeople,
                    emptyLabel,
                  );
                }
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
                        (!spRole && person.permissionRole === "Manager")
                      );
                    }
                    if (role === "Customer") {
                      return (
                        spRole === "supervisor" ||
                        (!spRole && person.permissionRole === "Supervisor")
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
                  options = withAssignedWorkforcePersonOption(
                    peopleOptions,
                    String(form[field.name] ?? ""),
                    livePermissionPeople,
                    emptyLabel,
                  );
                }
              } else if (field.companyScopedDepartments === "all") {
                const valueMode = field.departmentValueMode ?? "name";
                const allDepts = departments;
                options =
                  allDepts.length === 0
                    ? [
                        {
                          value: "",
                          label:
                            "No departments — add them under Departments",
                        },
                      ]
                    : [
                        { value: "", label: "— None —" },
                        ...allDepts.map((dept) => ({
                          value: valueMode === "id" ? dept.id : dept.name,
                          label: dept.companyName
                            ? `${dept.name} — ${dept.companyName}`
                            : dept.name,
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
                      idsEqual(dept.companyId, selectedCompanyId)
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

              // Quick-add button is shown next to any strict-RoleType dropdown
              // unless the field explicitly opts out.
              const quickAddEnabled =
                Boolean(field.sharePointRoleTypeFilter) &&
                field.allowQuickAdd !== false;
              const quickAddCompanyId =
                selectedCompanyId ||
                companies.find(
                  (company) =>
                    company.companyName.trim().toLowerCase() ===
                    selectedCompanyName,
                )?.id ||
                "";
              const quickAddCompanyName =
                String(form.companyName ?? "").trim() ||
                findCompanyById(companies, selectedCompanyId)?.companyName ||
                "";
              const selectValue = isSelectNone(String(form[field.name] ?? ""))
                ? field.permissionRoleFilter || field.sharePointRoleTypeFilter
                  ? String(form[field.name] ?? "")
                  : SELECT_NONE
                : String(form[field.name] ?? "");
              control = (
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>{field.label}</span>
                  <select
                    key={`${field.name}-${editing?.id ?? "new"}-${drawerOpen ? "open" : "closed"}`}
                    className={styles.select}
                    value={selectValue}
                    disabled={false}
                    aria-disabled={field.readOnly ? true : undefined}
                    style={
                      field.readOnly
                        ? { pointerEvents: "none", opacity: 0.85 }
                        : undefined
                    }
                    onChange={(event) =>
                      commitSelectValue(field, event.target.value)
                    }
                    onInput={(event) =>
                      commitSelectValue(
                        field,
                        (event.target as HTMLSelectElement).value,
                      )
                    }
                  >
                    {field.permissionRoleFilter ||
                    field.sharePointRoleTypeFilter ? null : (
                      <option value={SELECT_NONE}>Select…</option>
                    )}
                    {options.map((option, index) => (
                      <option
                        key={`${field.name}-${option.value}-${index}`}
                        value={option.value}
                      >
                        {option.label}
                      </option>
                    ))}
                  </select>
                  {field.type === "company" && field.name === "companyId" ? (
                    (() => {
                      const currentId = String(form.companyId ?? "").trim();
                      const currentLabel = !isSelectNone(currentId)
                        ? findCompanyById(companies, currentId)?.companyName
                        : null;
                      return (
                        <span
                          className={styles.fieldConfirmPill}
                          data-empty={currentLabel ? "false" : "true"}
                          aria-live="polite"
                        >
                          {currentLabel
                            ? `Selected company: ${currentLabel}`
                            : "No company selected yet"}
                        </span>
                      );
                    })()
                  ) : null}
                  {quickAddEnabled && quickAddCompanyId && quickAddCompanyName ? (
                    <button
                      type="button"
                      className={styles.quickAddInlineButton}
                      onClick={() => {
                        setQuickAdd({
                          fieldName: field.name,
                          role:
                            field.sharePointRoleTypeFilter as
                              | "Training Manager"
                              | "Supervisor"
                              | "Candidate",
                          companyId: quickAddCompanyId,
                          companyName: quickAddCompanyName,
                        });
                      }}
                      disabled={field.readOnly}
                    >
                      + Add new {field.sharePointRoleTypeFilter}
                    </button>
                  ) : null}
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
                          : "text"
                    }
                    inputMode={field.type === "email" ? "email" : undefined}
                    autoComplete={
                      field.type === "email" ? "email" : undefined
                    }
                    value={String(form[field.name] ?? "")}
                    placeholder={field.placeholder}
                    readOnly={field.readOnly}
                    disabled={field.readOnly}
                    onChange={(event) => {
                      const value = event.target.value;
                      setForm((current) => {
                        const next: FormState = {
                          ...current,
                          [field.name]: value,
                        };
                        if (
                          field.name === "trainingDate" &&
                          shouldAutoFillBlankExpiry(
                            fields,
                            String(current.trainingOutcome ?? ""),
                          ) &&
                          !String(current.expiry ?? "").trim()
                        ) {
                          const filled = autoExpiryIso(fields, value);
                          if (filled) next.expiry = filled;
                        }
                        return next;
                      });
                    }}
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

      <QuickAddPermissionPersonModal
        open={quickAdd !== null}
        role={quickAdd?.role ?? "Training Manager"}
        companyId={quickAdd?.companyId ?? ""}
        companyName={quickAdd?.companyName ?? ""}
        onClose={() => setQuickAdd(null)}
        onCreated={(record: AdminPermissionRecord) => {
          // Show the new person in the dropdown immediately (don't wait on
          // the list refresh), then refresh in the background for consistency.
          setPermissionPeopleLive((current) => {
            const next = current ?? permissionPeople;
            const entry: AdminPermissionPersonOption = {
              id: record.id,
              userEmail: record.userEmail,
              name: record.name ?? null,
              status: record.status,
              permissionRole: record.permissionRole,
              sharePointRoleType: record.sharePointRoleType ?? null,
              companyId: record.companyId ?? null,
              companyName: record.companyName ?? null,
            };
            if (next.some((row) => row.id === entry.id)) {
              return next.map((row) => (row.id === entry.id ? entry : row));
            }
            return [...next, entry];
          });
          void refreshPermissionPeople();
          // Auto-select the newly created person on the field that opened
          // this modal. The dropdown value is the Permissions row's Name
          // (or userEmail fallback) — matches what applyWorkforcePersonLookups
          // resolves back to a Lookup id.
          const selectValue = record.name?.trim() || record.userEmail;
          if (quickAdd) {
            setForm((current) => {
              const next: FormState = {
                ...current,
                [quickAdd.fieldName]: selectValue,
              };
              // Email helps Workforce save resolve the Permissions row when
              // the display name collides or refresh is still in flight.
              if (quickAdd.fieldName === "trainingManager") {
                next.trainingManagerEmail = record.userEmail;
              } else if (quickAdd.fieldName === "supervisor") {
                next.supervisorEmail = record.userEmail;
              }
              return next;
            });
          }
          pushToast(`Added ${record.name ?? record.userEmail}`, "success");
          setQuickAdd(null);
        }}
      />
    </div>
  );
}

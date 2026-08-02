import * as React from "react";

import type { SharePointListKey } from "../schema/sharepointSchema";
import {
  getSharePointList,
  SHAREPOINT_LISTS,
} from "../schema/sharepointSchema";
import {
  thumbnailPreviewUrl,
  uploadAndSetListImage,
} from "../services/companyLogoService";
import {
  exportTableAsCsv,
  exportTableAsExcel,
} from "../services/exportService";
import type { PortalTableRow } from "../services/portalDataService";
import {
  asString,
  createListItem,
  deleteListItem,
  getListItem,
  updateListItem,
  type SpListClient,
} from "../services/sharePointListService";
import { formatDate, formatDateTime } from "../utils/formatDate";
import styles from "./portal.module.scss";

export interface AdminDataTableProps {
  client: SpListClient;
  listKey: SharePointListKey;
  title: string;
  headers: string[];
  columns: string[];
  rows: PortalTableRow[];
  loading: boolean;
  error: string | null;
  onRefresh: () => void;
  /** Optional: prefer these columns in the edit form (schema order). */
  formColumns?: string[];
}

const READONLY_FIELDS: { [key: string]: boolean } = {
  ID: true,
  Id: true,
  FileLeafRef: true,
  FileRef: true,
  FileDirRef: true,
  UniqueId: true,
  GUID: true,
  Created: true,
  Modified: true,
  AuthorId: true,
  EditorId: true,
  ContentTypeId: true,
  Attachments: true,
  CompanyLookupId: true,
  CompanyLogo: true,
};

const BOOLEAN_HINTS = [
  "CustomerVisible",
  "Customer_x0020_Visible",
  "CanView",
  "CanDownload",
  "CanEdit",
  "Active",
  "DoNotSync",
  "ReceiveExpiryNotifications",
  "ReceiveDocumentNotifications",
  "CustomerNotificationsEnabled",
  "NeedsReview",
];

/** Choice fields — dropdowns instead of free text (matches Next.js admin). */
const CHOICE_OPTIONS: { [internalName: string]: string[] } = {
  RoleType: ["Admin", "Customer", "Candidate"],
  AccessScope: ["Full Company", "Department Only", "Candidate Only"],
  Status: ["Active", "Inactive"],
  TrainingOutcome: ["Pass", "Fail"],
  CompanySize: ["Small", "Medium", "Large", "Enterprise"],
  OverallStatus: ["Compliant", "Expiring Soon", "Expired", "Missing Data"],
  SyncStatus: ["Synced", "Pending", "Error", "Skipped"],
  Category: ["Promotion", "Offer", "Announcement"],
  NoviceorEwt: ["Novice", "EWT"],
};

const DATE_HINTS = [
  /Date$/i,
  /Expiry/i,
  /^DOB$/i,
  /EventDate/i,
  /StartDate/i,
  /EndDate/i,
  /OutcomeDate/i,
  /TrainingDate/i,
  /CourseDate/i,
  /Timestamp/i,
  /LastSynced/i,
];

function isBooleanField(internalName: string): boolean {
  return BOOLEAN_HINTS.indexOf(internalName) >= 0;
}

function isLookupIdField(internalName: string): boolean {
  return (
    /Id$/i.test(internalName) &&
    internalName !== "ID" &&
    internalName !== "Id"
  );
}

function isDateField(internalName: string): boolean {
  for (let i = 0; i < DATE_HINTS.length; i++) {
    if (DATE_HINTS[i].test(internalName)) return true;
  }
  return false;
}

function friendlyLabel(
  listKey: SharePointListKey,
  internalName: string
): string {
  const list = SHAREPOINT_LISTS[listKey];
  const fields = list.fields as Record<string, string>;
  const labels = list.labels as Record<string, string>;
  for (const key in fields) {
    if (
      Object.prototype.hasOwnProperty.call(fields, key) &&
      fields[key] === internalName
    ) {
      return labels[key] || humanize(internalName);
    }
  }
  return humanize(internalName);
}

function humanize(internalName: string): string {
  return internalName
    .replace(/_x0020_/g, " ")
    .replace(/_x002d_/g, " - ")
    .replace(/_x002f_/g, " / ")
    .replace(/_x2013_/g, " – ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\s+/g, " ")
    .trim();
}

function fieldToFormValue(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "object") {
    const s = asString(value);
    return s || "";
  }
  const raw = String(value);
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
  return raw;
}

function isDateTimeDisplayField(
  listKey: SharePointListKey,
  internalName: string
): boolean {
  if (listKey === "events") {
    return ["EventDate", "EndDate", "LastSyncedAt"].indexOf(internalName) >= 0;
  }
  return listKey === "trainingManagerLogs" && internalName === "Timestamp";
}

function formatCellDisplay(
  value: string,
  listKey: SharePointListKey,
  internalName: string
): string {
  if (!value) return "—";
  if (isDateTimeDisplayField(listKey, internalName)) {
    return formatDateTime(value);
  }
  if (isDateField(internalName) || /^\d{4}-\d{2}-\d{2}(?:T|$)/.test(value)) {
    return formatDate(value);
  }
  return value;
}

function buildWritePayload(
  columns: string[],
  form: Record<string, string>,
  original: Record<string, unknown>
): Record<string, unknown> {
  const payload: Record<string, unknown> = {};

  for (let i = 0; i < columns.length; i++) {
    const name = columns[i];
    if (READONLY_FIELDS[name]) continue;
    if (!Object.prototype.hasOwnProperty.call(form, name)) continue;

    const idCompanion = name + "Id";

    if (isLookupIdField(name)) {
      const raw = (form[name] || "").trim();
      payload[name] =
        raw === "" ? null : /^\d+$/.test(raw) ? Number(raw) : raw;
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(original, idCompanion)) {
      const raw = (form[name] || "").trim();
      payload[idCompanion] =
        raw === "" ? null : /^\d+$/.test(raw) ? Number(raw) : null;
      continue;
    }

    if (isBooleanField(name)) {
      const raw = (form[name] || "").trim().toLowerCase();
      if (raw === "") continue;
      payload[name] =
        raw === "true" || raw === "1" || raw === "yes" || raw === "on";
      continue;
    }

    const text = (form[name] || "").trim();
    if (text === "") continue;
    if (isDateField(name) && /^\d{4}-\d{2}-\d{2}$/.test(text)) {
      payload[name] = text + "T00:00:00Z";
      continue;
    }
    payload[name] = text;
  }

  return payload;
}

export const AdminDataTable: React.FC<AdminDataTableProps> = (props) => {
  const {
    client,
    listKey,
    title,
    headers,
    columns,
    rows,
    loading,
    error,
    onRefresh,
    formColumns,
  } = props;

  const [drawerOpen, setDrawerOpen] = React.useState(false);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [form, setForm] = React.useState<Record<string, string>>({});
  const [original, setOriginal] = React.useState<Record<string, unknown>>({});
  const [busy, setBusy] = React.useState(false);
  const [actionError, setActionError] = React.useState<string | null>(null);
  const [actionOk, setActionOk] = React.useState<string | null>(null);
  const [logoFile, setLogoFile] = React.useState<File | null>(null);
  const [logoPreview, setLogoPreview] = React.useState<string | null>(null);
  const [localPreviewUrl, setLocalPreviewUrl] = React.useState<string | null>(
    null
  );
  const [query, setQuery] = React.useState("");

  const listLabel = getSharePointList(listKey).displayName;
  const supportsCompanyLogo = listKey === "company";

  const editableColumns = React.useMemo(() => {
    const source = formColumns && formColumns.length ? formColumns : columns;
    return source.filter((c) => !READONLY_FIELDS[c]);
  }, [columns, formColumns]);

  const filteredRows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return rows;
    return rows.filter((row) =>
      row.cells.some((c) => (c || "").toLowerCase().indexOf(q) >= 0)
    );
  }, [rows, query]);

  const clearLogoSelection = (): void => {
    if (localPreviewUrl) {
      URL.revokeObjectURL(localPreviewUrl);
    }
    setLocalPreviewUrl(null);
    setLogoFile(null);
  };

  const openCreate = (): void => {
    const blank: Record<string, string> = {};
    for (let i = 0; i < editableColumns.length; i++) {
      const name = editableColumns[i];
      if (name === "Status") blank[name] = "Active";
      else if (name === "RoleType") blank[name] = "Customer";
      else if (name === "AccessScope") blank[name] = "Full Company";
      else if (name === "TrainingOutcome") blank[name] = "Pass";
      else if (isBooleanField(name)) blank[name] = "true";
      else blank[name] = "";
    }
    setEditingId(null);
    setOriginal({});
    setForm(blank);
    clearLogoSelection();
    setLogoPreview(null);
    setActionError(null);
    setActionOk(null);
    setDrawerOpen(true);
  };

  const openEdit = async (row: PortalTableRow): Promise<void> => {
    setBusy(true);
    setActionError(null);
    setActionOk(null);
    try {
      const item = await getListItem(client, listKey, row.id);
      const fields = item ? item.fields : row.fields || {};
      const next: Record<string, string> = {};
      for (let i = 0; i < editableColumns.length; i++) {
        const name = editableColumns[i];
        const idKey = name + "Id";
        if (fields[idKey] !== undefined && fields[idKey] !== null) {
          next[name] = fieldToFormValue(fields[idKey]);
        } else {
          next[name] = fieldToFormValue(fields[name]);
        }
      }
      for (const key in fields) {
        if (
          Object.prototype.hasOwnProperty.call(fields, key) &&
          isLookupIdField(key) &&
          !READONLY_FIELDS[key] &&
          next[key] === undefined
        ) {
          next[key] = fieldToFormValue(fields[key]);
        }
      }
      setEditingId(row.id);
      setOriginal(fields);
      setForm(next);
      clearLogoSelection();
      setLogoPreview(thumbnailPreviewUrl(client.webUrl, fields.CompanyLogo));
      setDrawerOpen(true);
    } catch (e) {
      setActionError(e instanceof Error ? e.message : "Failed to load item.");
    } finally {
      setBusy(false);
    }
  };

  const onLogoChosen = (fileList: FileList | null): void => {
    const file = fileList && fileList.length > 0 ? fileList[0] : null;
    clearLogoSelection();
    if (!file) return;
    if (file.type && file.type.indexOf("image/") !== 0) {
      setActionError("Please choose an image file (PNG, JPG, WEBP, etc.).");
      return;
    }
    setActionError(null);
    setLogoFile(file);
    setLocalPreviewUrl(URL.createObjectURL(file));
  };

  const save = async (): Promise<void> => {
    setBusy(true);
    setActionError(null);
    setActionOk(null);
    try {
      const payload = buildWritePayload(editableColumns, form, original);
      for (const key in form) {
        if (
          Object.prototype.hasOwnProperty.call(form, key) &&
          isLookupIdField(key) &&
          !READONLY_FIELDS[key]
        ) {
          const raw = (form[key] || "").trim();
          if (raw === "") {
            if (editingId) payload[key] = null;
          } else {
            payload[key] = /^\d+$/.test(raw) ? Number(raw) : raw;
          }
        }
      }

      let itemId = editingId;
      if (editingId) {
        await updateListItem(client, listKey, editingId, payload);
      } else {
        const created = await createListItem(client, listKey, payload);
        itemId = created.id;
      }

      if (supportsCompanyLogo && logoFile && itemId) {
        await uploadAndSetListImage(
          client,
          listKey,
          itemId,
          "CompanyLogo",
          logoFile
        );
      }

      setActionOk(
        editingId
          ? "Saved item #" + editingId + (logoFile ? " (logo updated)" : "")
          : "Created item #" + itemId + (logoFile ? " (logo uploaded)" : "")
      );
      setDrawerOpen(false);
      clearLogoSelection();
      onRefresh();
    } catch (e) {
      setActionError(e instanceof Error ? e.message : "Save failed.");
    } finally {
      setBusy(false);
    }
  };

  const remove = async (row: PortalTableRow): Promise<void> => {
    const ok = window.confirm(
      "Delete item #" +
        row.id +
        " from " +
        listLabel +
        "? This cannot be undone."
    );
    if (!ok) return;
    setBusy(true);
    setActionError(null);
    setActionOk(null);
    try {
      await deleteListItem(client, listKey, row.id);
      setActionOk("Deleted item #" + row.id);
      onRefresh();
    } catch (e) {
      setActionError(e instanceof Error ? e.message : "Delete failed.");
    } finally {
      setBusy(false);
    }
  };

  const exportRows = filteredRows.map((r) => r.cells);
  const formKeys = editableColumns.filter(
    (k) => Object.prototype.hasOwnProperty.call(form, k) || true
  );
  const previewSrc = localPreviewUrl || logoPreview;

  const renderFieldControl = (key: string): React.ReactNode => {
    const choices = CHOICE_OPTIONS[key];
    if (choices) {
      return (
        <select
          value={form[key] || ""}
          onChange={(e) => setForm({ ...form, [key]: e.target.value })}
          disabled={busy}
        >
          <option value="">Select…</option>
          {choices.map((opt) => (
            <option key={opt} value={opt}>
              {opt}
            </option>
          ))}
        </select>
      );
    }
    if (isBooleanField(key)) {
      return (
        <select
          value={form[key] || "false"}
          onChange={(e) => setForm({ ...form, [key]: e.target.value })}
          disabled={busy}
        >
          <option value="true">Yes</option>
          <option value="false">No</option>
        </select>
      );
    }
    if (isDateField(key)) {
      return (
        <input
          type="date"
          value={(form[key] || "").slice(0, 10)}
          onChange={(e) => setForm({ ...form, [key]: e.target.value })}
          disabled={busy}
        />
      );
    }
    if (key === "UserEmail" || /Email$/i.test(key)) {
      return (
        <input
          type="email"
          value={form[key] || ""}
          onChange={(e) => setForm({ ...form, [key]: e.target.value })}
          disabled={busy}
          placeholder="name@company.com"
        />
      );
    }
    return (
      <input
        type="text"
        value={form[key] || ""}
        onChange={(e) => setForm({ ...form, [key]: e.target.value })}
        disabled={busy}
        placeholder={
          isLookupIdField(key) || original[key + "Id"] !== undefined
            ? "SharePoint item ID"
            : undefined
        }
      />
    );
  };

  return (
    <div className={styles.panel}>
      <div className={styles.toolbar}>
        <button type="button" onClick={openCreate} disabled={busy}>
          Add new
        </button>
        <button
          type="button"
          onClick={() => exportTableAsCsv(title, headers, exportRows)}
          disabled={loading || filteredRows.length === 0}
        >
          Export CSV
        </button>
        <button
          type="button"
          onClick={() => exportTableAsExcel(title, headers, exportRows)}
          disabled={loading || filteredRows.length === 0}
        >
          Export Excel
        </button>
        <button type="button" onClick={onRefresh} disabled={busy || loading}>
          Refresh
        </button>
        <input
          className={styles.tableSearch}
          type="search"
          placeholder="Search rows…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          aria-label="Search table"
        />
      </div>

      {actionError && <p className={styles.error}>{actionError}</p>}
      {actionOk && <p className={styles.success}>{actionOk}</p>}
      {loading && <p className={styles.muted}>Loading…</p>}
      {error && <p className={styles.error}>{error}</p>}

      {!loading && !error && (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th className={styles.actionsCol}>Actions</th>
                {headers.map((h) => (
                  <th key={h} title={h}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredRows.length === 0 ? (
                <tr>
                  <td
                    colSpan={(headers.length || 0) + 1}
                    className={styles.muted}
                  >
                    No rows found.
                  </td>
                </tr>
              ) : (
                filteredRows.map((row) => (
                  <tr key={row.id}>
                    <td className={styles.actionsCol}>
                      <button
                        type="button"
                        className={styles.linkBtn}
                        onClick={() => {
                          openEdit(row).catch(() => undefined);
                        }}
                        disabled={busy}
                      >
                        Edit
                      </button>
                      <button
                        type="button"
                        className={styles.linkBtnDanger}
                        onClick={() => {
                          remove(row).catch(() => undefined);
                        }}
                        disabled={busy}
                      >
                        Delete
                      </button>
                    </td>
                    {row.cells.map((c, i) => {
                      const display = formatCellDisplay(
                        c,
                        listKey,
                        columns[i] || ""
                      );
                      return (
                        <td key={row.id + "-" + i} title={display}>
                          {display}
                        </td>
                      );
                    })}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      )}

      {drawerOpen && (
        <div
          className={styles.drawerBackdrop}
          onClick={() => {
            if (!busy) {
              clearLogoSelection();
              setDrawerOpen(false);
            }
          }}
        >
          <div
            className={styles.drawer}
            role="dialog"
            aria-modal="true"
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.drawerHeader}>
              <h2 className={styles.title}>
                {editingId ? "Edit" : "Add"} {listLabel}
                {editingId ? " #" + editingId : ""}
              </h2>
              <button
                type="button"
                className={styles.linkBtn}
                onClick={() => {
                  clearLogoSelection();
                  setDrawerOpen(false);
                }}
                disabled={busy}
              >
                Close
              </button>
            </div>
            <p className={styles.muted}>
              Use the fields below. Choice fields use dropdowns. Blank fields
              are skipped on save.
            </p>
            {actionError && <p className={styles.error}>{actionError}</p>}

            {supportsCompanyLogo && (
              <div className={styles.logoBox}>
                <span className={styles.fieldLabel}>Company Logo</span>
                {previewSrc ? (
                  <img
                    src={previewSrc}
                    alt="Company logo preview"
                    className={styles.logoPreview}
                  />
                ) : (
                  <p className={styles.muted}>No logo yet</p>
                )}
                <input
                  type="file"
                  accept="image/*"
                  disabled={busy}
                  onChange={(e) => onLogoChosen(e.target.files)}
                />
                {logoFile && (
                  <p className={styles.muted}>
                    Selected: {logoFile.name} — will upload on Save
                  </p>
                )}
                {logoFile && (
                  <button
                    type="button"
                    className={styles.linkBtn}
                    disabled={busy}
                    onClick={clearLogoSelection}
                  >
                    Clear selection
                  </button>
                )}
              </div>
            )}

            <div className={styles.formGrid}>
              {formKeys.map((key) => (
                <label key={key} className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {friendlyLabel(listKey, key)}
                    {CHOICE_OPTIONS[key] ? "" : ""}
                    {isLookupIdField(key) ||
                    original[key + "Id"] !== undefined
                      ? " (lookup ID)"
                      : ""}
                  </span>
                  {renderFieldControl(key)}
                </label>
              ))}
            </div>
            <div className={styles.drawerFooter}>
              <button type="button" onClick={save} disabled={busy}>
                {busy
                  ? "Saving…"
                  : editingId
                    ? "Save changes"
                    : "Create"}
              </button>
              <button
                type="button"
                onClick={() => {
                  clearLogoSelection();
                  setDrawerOpen(false);
                }}
                disabled={busy}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

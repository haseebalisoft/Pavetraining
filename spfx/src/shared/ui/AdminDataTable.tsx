import * as React from "react";

import type { SharePointListKey } from "../schema/sharepointSchema";
import { getSharePointList } from "../schema/sharepointSchema";
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

function fieldToFormValue(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "object") {
    const s = asString(value);
    return s || "";
  }
  return String(value);
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

  const listLabel = getSharePointList(listKey).displayName;
  const supportsCompanyLogo = listKey === "company";

  const clearLogoSelection = (): void => {
    if (localPreviewUrl) {
      URL.revokeObjectURL(localPreviewUrl);
    }
    setLocalPreviewUrl(null);
    setLogoFile(null);
  };

  const openCreate = (): void => {
    const blank: Record<string, string> = {};
    for (let i = 0; i < columns.length; i++) {
      if (!READONLY_FIELDS[columns[i]]) blank[columns[i]] = "";
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
      for (let i = 0; i < columns.length; i++) {
        const name = columns[i];
        if (READONLY_FIELDS[name]) continue;
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
          !READONLY_FIELDS[key]
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
      const payload = buildWritePayload(columns, form, original);
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

  const exportRows = rows.map((r) => r.cells);
  const formKeys = Object.keys(form).sort();
  const previewSrc = localPreviewUrl || logoPreview;

  return (
    <div className={styles.panel}>
      <div className={styles.toolbar}>
        <button type="button" onClick={openCreate} disabled={busy}>
          New
        </button>
        <button
          type="button"
          onClick={() => exportTableAsCsv(title, headers, exportRows)}
          disabled={loading || rows.length === 0}
        >
          Export CSV
        </button>
        <button
          type="button"
          onClick={() => exportTableAsExcel(title, headers, exportRows)}
          disabled={loading || rows.length === 0}
        >
          Export Excel
        </button>
        <button type="button" onClick={onRefresh} disabled={busy || loading}>
          Refresh
        </button>
      </div>

      {actionError && <p className={styles.error}>{actionError}</p>}
      {actionOk && <p className={styles.success}>{actionOk}</p>}
      {loading && <p className={styles.muted}>Loading…</p>}
      {error && <p className={styles.error}>{error}</p>}

      {!loading && !error && (
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.actionsCol}>Actions</th>
              {headers.map((h) => (
                <th key={h}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td
                  colSpan={(headers.length || 0) + 1}
                  className={styles.muted}
                >
                  No rows found.
                </td>
              </tr>
            ) : (
              rows.map((row) => (
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
                  {row.cells.map((c, i) => (
                    <td key={row.id + "-" + i}>{c}</td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      )}

      {drawerOpen && (
        <div className={styles.drawerBackdrop}>
          <div className={styles.drawer} role="dialog" aria-modal="true">
            <div className={styles.drawerHeader}>
              <h2 className={styles.title}>
                {editingId ? "Edit item #" + editingId : "New item"} —{" "}
                {listLabel}
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
              Fill Company Name (required). Blank fields are skipped. Lookup
              columns: enter the SharePoint item ID.
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
                    {key}
                    {isBooleanField(key) ? " (yes/no)" : ""}
                    {original[key + "Id"] !== undefined
                      ? " (lookup ID)"
                      : ""}
                  </span>
                  {isBooleanField(key) ? (
                    <select
                      value={form[key] || "false"}
                      onChange={(e) =>
                        setForm({ ...form, [key]: e.target.value })
                      }
                      disabled={busy}
                    >
                      <option value="true">Yes</option>
                      <option value="false">No</option>
                    </select>
                  ) : (
                    <input
                      type="text"
                      value={form[key] || ""}
                      onChange={(e) =>
                        setForm({ ...form, [key]: e.target.value })
                      }
                      disabled={busy}
                    />
                  )}
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

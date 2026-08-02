"use client";

import { StatusBadge } from "@/components/ui/StatusBadge";
import type { AdminDocumentRecord } from "@/types/adminDocuments";
import { formatDate } from "@/lib/utils/formatDate";

import { DocumentActionsMenu } from "./DocumentActionsMenu";
import styles from "./documentsBrowse.module.css";

function yesNo(value: boolean) {
  return (
    <StatusBadge
      label={value ? "Yes" : "No"}
      tone={value ? "ok" : "neutral"}
    />
  );
}

/**
 * SharePoint-aligned Customer Documents columns:
 * Name, Modified, Modified By, ID, Company, Candidate, Document Type,
 * Customer Visible, Notification Sent, Notify Customer.
 */
export function DocumentsTableView({
  rows,
  selectedIds,
  busyId,
  onToggleSelect,
  onToggleSelectAll,
  onOpenFolder,
  onEditMetadata,
  onSetVisibility,
}: {
  rows: AdminDocumentRecord[];
  selectedIds: Set<string>;
  busyId: string | null;
  onToggleSelect: (id: string) => void;
  onToggleSelectAll: (selectAll: boolean) => void;
  onOpenFolder?: (row: AdminDocumentRecord) => void;
  onEditMetadata: (row: AdminDocumentRecord) => void;
  onSetVisibility: (
    row: AdminDocumentRecord,
    customerVisible: boolean,
  ) => void;
}) {
  if (rows.length === 0) {
    return (
      <div className={styles.emptyState}>
        <h2>This folder is empty</h2>
        <p>Upload a file or open another folder in Customer Documents.</p>
      </div>
    );
  }

  const selectable = rows.filter((row) => !row.isFolder);
  const allSelected =
    selectable.length > 0 && selectable.every((row) => selectedIds.has(row.id));

  return (
    <div className={styles.tableWrap}>
      <table className={styles.dataTable}>
        <thead>
          <tr>
            <th scope="col" className={styles.selectCol}>
              <input
                type="checkbox"
                aria-label="Select all files"
                checked={allSelected}
                onChange={(event) => onToggleSelectAll(event.target.checked)}
              />
            </th>
            <th scope="col">Name</th>
            <th scope="col">Modified</th>
            <th scope="col">Modified By</th>
            <th scope="col">ID</th>
            <th scope="col">Company</th>
            <th scope="col">Candidate</th>
            <th scope="col">Document Type</th>
            <th scope="col">Customer Visible</th>
            <th scope="col">Notification Sent</th>
            <th scope="col">Notify Customer</th>
            <th scope="col">Actions</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const selected = selectedIds.has(row.id);
            return (
              <tr
                key={row.id}
                className={row.isFolder ? styles.folderRow : undefined}
              >
                <td className={styles.selectCol}>
                  {row.isFolder ? null : (
                    <input
                      type="checkbox"
                      aria-label={`Select ${row.name}`}
                      checked={selected}
                      onChange={() => onToggleSelect(row.id)}
                    />
                  )}
                </td>
                <td>
                  {row.isFolder ? (
                    <button
                      type="button"
                      className={styles.folderNameButton}
                      onClick={() => onOpenFolder?.(row)}
                    >
                      <span className={styles.folderIcon} aria-hidden="true">
                        <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                          <path d="M3.5 6.75A2.25 2.25 0 0 1 5.75 4.5h4.1c.45 0 .88.18 1.2.5l1.2 1.2c.32.32.75.5 1.2.5h5c1.24 0 2.25 1.01 2.25 2.25v8.3A2.25 2.25 0 0 1 17.45 19.5H5.75A2.25 2.25 0 0 1 3.5 17.25V6.75Z" />
                        </svg>
                      </span>
                      <strong>{row.name}</strong>
                    </button>
                  ) : (
                    <strong>{row.name}</strong>
                  )}
                </td>
                <td title={formatDate(row.modifiedDate)}>
                  {formatDate(row.modifiedDate)}
                </td>
                <td>{row.modifiedBy?.trim() || "—"}</td>
                <td>{row.id}</td>
                <td>{row.company?.trim() || "—"}</td>
                <td>{row.candidate?.trim() || "—"}</td>
                <td>{row.documentType?.trim() || "—"}</td>
                <td>{yesNo(row.customerVisible)}</td>
                <td>{yesNo(row.notificationSent)}</td>
                <td>{yesNo(row.notifyCustomer)}</td>
                <td>
                  {row.isFolder ? (
                    <button
                      type="button"
                      className={styles.menuTrigger}
                      onClick={() => onOpenFolder?.(row)}
                    >
                      Open
                    </button>
                  ) : (
                    <DocumentActionsMenu
                      row={row}
                      busy={busyId === row.id}
                      onEditMetadata={() => onEditMetadata(row)}
                      onSetVisibility={(visible) =>
                        onSetVisibility(row, visible)
                      }
                    />
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

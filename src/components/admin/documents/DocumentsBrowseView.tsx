"use client";

import { AnimatePresence, motion } from "framer-motion";

import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type {
  AdminDocumentRecord,
  DocumentMetadataStatus,
} from "@/lib/services/adminCrudService";

import { DocumentActionsMenu } from "./DocumentActionsMenu";
import { FolderGrid, type FolderItem } from "./FolderGrid";
import styles from "./documentsBrowse.module.css";

export type BrowsePath =
  | { level: "companies" }
  | { level: "types"; companyKey: string; companyLabel: string }
  | {
      level: "files";
      companyKey: string;
      companyLabel: string;
      typeKey: string;
      typeLabel: string;
    };

export const UNASSIGNED_COMPANY_KEY = "__unassigned_company__";
export const UNASSIGNED_TYPE_KEY = "__unassigned_type__";

const slideVariants = {
  enter: (direction: number) => ({
    x: direction > 0 ? 28 : -28,
    opacity: 0,
  }),
  center: { x: 0, opacity: 1 },
  exit: (direction: number) => ({
    x: direction > 0 ? -28 : 28,
    opacity: 0,
  }),
};

function metadataStatusClass(status: DocumentMetadataStatus): string {
  if (status === "Complete") return styles.metadataComplete;
  if (status === "Hidden from Customer") return styles.metadataHidden;
  return styles.metadataMissing;
}

function FolderEmpty({
  title,
  message,
}: {
  title: string;
  message: string;
}) {
  return (
    <motion.div
      className={styles.emptyBrowse}
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      transition={{ duration: 0.25 }}
    >
      <motion.span
        className={styles.emptyIcon}
        aria-hidden="true"
        animate={{ y: [0, -5, 0] }}
        transition={{ duration: 2.4, repeat: Infinity, ease: "easeInOut" }}
      >
        <svg viewBox="0 0 24 24" fill="currentColor">
          <path d="M3.5 6.75A2.25 2.25 0 0 1 5.75 4.5h4.1c.45 0 .88.18 1.2.5l1.2 1.2c.32.32.75.5 1.2.5h5c1.24 0 2.25 1.01 2.25 2.25v8.3A2.25 2.25 0 0 1 17.45 19.5H5.75A2.25 2.25 0 0 1 3.5 17.25V6.75Z" />
        </svg>
      </motion.span>
      <h2>{title}</h2>
      <p>{message}</p>
    </motion.div>
  );
}

function BrowseCrumbs({
  path,
  onNavigate,
}: {
  path: BrowsePath;
  onNavigate: (next: BrowsePath, direction: -1 | 1) => void;
}) {
  return (
    <nav className={styles.browseCrumbs} aria-label="Folder location">
      <button
        type="button"
        className={
          path.level === "companies" ? styles.crumbCurrent : styles.crumbButton
        }
        disabled={path.level === "companies"}
        onClick={() => onNavigate({ level: "companies" }, -1)}
      >
        Companies
      </button>
      {path.level !== "companies" ? (
        <>
          <span className={styles.crumbSep} aria-hidden="true">
            /
          </span>
          <button
            type="button"
            className={
              path.level === "types" ? styles.crumbCurrent : styles.crumbButton
            }
            disabled={path.level === "types"}
            onClick={() =>
              onNavigate(
                {
                  level: "types",
                  companyKey: path.companyKey,
                  companyLabel: path.companyLabel,
                },
                -1,
              )
            }
          >
            {path.companyLabel}
          </button>
        </>
      ) : null}
      {path.level === "files" ? (
        <>
          <span className={styles.crumbSep} aria-hidden="true">
            /
          </span>
          <span className={styles.crumbCurrent}>{path.typeLabel}</span>
        </>
      ) : null}
    </nav>
  );
}

export function companyKeyFromRecord(row: AdminDocumentRecord): string {
  const value = row.company?.trim();
  return value ? value.toLowerCase() : UNASSIGNED_COMPANY_KEY;
}

export function companyLabelFromRecord(row: AdminDocumentRecord): string {
  return row.company?.trim() || "Unassigned company";
}

export function typeKeyFromRecord(row: AdminDocumentRecord): string {
  const value = row.documentType?.trim();
  return value ? value.toLowerCase() : UNASSIGNED_TYPE_KEY;
}

export function typeLabelFromRecord(row: AdminDocumentRecord): string {
  return row.documentType?.trim() || "Unassigned type";
}

export function DocumentsBrowseView({
  rows,
  path,
  direction,
  busyId,
  onNavigate,
  onEditMetadata,
  onSetVisibility,
}: {
  rows: AdminDocumentRecord[];
  path: BrowsePath;
  direction: number;
  busyId: string | null;
  onNavigate: (next: BrowsePath, direction: -1 | 1) => void;
  onEditMetadata: (row: AdminDocumentRecord) => void;
  onSetVisibility: (
    row: AdminDocumentRecord,
    customerVisible: boolean,
  ) => void;
}) {
  const companyFolders: FolderItem[] = (() => {
    const map = new Map<string, FolderItem>();
    for (const row of rows) {
      const key = companyKeyFromRecord(row);
      const existing = map.get(key);
      if (existing) {
        existing.count += 1;
      } else {
        map.set(key, {
          key,
          name: companyLabelFromRecord(row),
          count: 1,
          countNoun: "documents",
        });
      }
    }
    return Array.from(map.values()).sort((a, b) =>
      a.name.localeCompare(b.name),
    );
  })();

  const typeFolders: FolderItem[] =
    path.level === "companies"
      ? []
      : (() => {
          const map = new Map<string, FolderItem>();
          for (const row of rows) {
            if (companyKeyFromRecord(row) !== path.companyKey) continue;
            const key = typeKeyFromRecord(row);
            const existing = map.get(key);
            if (existing) {
              existing.count += 1;
            } else {
              map.set(key, {
                key,
                name: typeLabelFromRecord(row),
                count: 1,
                countNoun: "files",
              });
            }
          }
          return Array.from(map.values()).sort((a, b) =>
            a.name.localeCompare(b.name),
          );
        })();

  const files =
    path.level === "files"
      ? rows
          .filter(
            (row) =>
              companyKeyFromRecord(row) === path.companyKey &&
              typeKeyFromRecord(row) === path.typeKey,
          )
          .sort((a, b) => a.name.localeCompare(b.name))
      : [];

  const stageKey =
    path.level === "companies"
      ? "companies"
      : path.level === "types"
        ? `types:${path.companyKey}`
        : `files:${path.companyKey}:${path.typeKey}`;

  let metaLabel = `${companyFolders.length} compan${companyFolders.length === 1 ? "y" : "ies"}`;
  if (path.level === "types") {
    metaLabel = `${typeFolders.length} document type${typeFolders.length === 1 ? "" : "s"}`;
  } else if (path.level === "files") {
    metaLabel = `${files.length} file${files.length === 1 ? "" : "s"}`;
  }

  return (
    <div className={styles.browseRoot}>
      <div className={styles.levelHeader}>
        <BrowseCrumbs path={path} onNavigate={onNavigate} />
        <p className={styles.levelMeta}>{metaLabel}</p>
      </div>

      <div className={styles.stage}>
        <AnimatePresence mode="wait" custom={direction}>
          <motion.div
            key={stageKey}
            custom={direction}
            variants={slideVariants}
            initial="enter"
            animate="center"
            exit="exit"
            transition={{ duration: 0.22, ease: "easeOut" }}
          >
            {path.level === "companies" ? (
              companyFolders.length === 0 ? (
                <FolderEmpty
                  title="No company folders"
                  message="No documents match the current filters."
                />
              ) : (
                <FolderGrid
                  items={companyFolders}
                  countNoun="documents"
                  onOpen={(key) => {
                    const folder = companyFolders.find((item) => item.key === key);
                    if (!folder) return;
                    onNavigate(
                      {
                        level: "types",
                        companyKey: folder.key,
                        companyLabel: folder.name,
                      },
                      1,
                    );
                  }}
                />
              )
            ) : null}

            {path.level === "types" ? (
              typeFolders.length === 0 ? (
                <FolderEmpty
                  title="No document types"
                  message="No document types found for this company with the current filters."
                />
              ) : (
                <FolderGrid
                  items={typeFolders}
                  countNoun="files"
                  onOpen={(key) => {
                    const folder = typeFolders.find((item) => item.key === key);
                    if (!folder) return;
                    onNavigate(
                      {
                        level: "files",
                        companyKey: path.companyKey,
                        companyLabel: path.companyLabel,
                        typeKey: folder.key,
                        typeLabel: folder.name,
                      },
                      1,
                    );
                  }}
                />
              )
            ) : null}

            {path.level === "files" ? (
              files.length === 0 ? (
                <FolderEmpty
                  title="No files"
                  message="No files in this folder match the current filters."
                />
              ) : (
                <div className={styles.fileList}>
                  {files.map((row, index) => (
                    <motion.article
                      key={row.id}
                      className={`${styles.fileRow} ${
                        row.metadataStatus === "Complete"
                          ? ""
                          : styles.fileRowNeedsMeta
                      }`}
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{
                        duration: 0.2,
                        delay: index * 0.045,
                        ease: "easeOut",
                      }}
                    >
                      <div className={styles.fileMain}>
                        <div className={styles.fileTitleRow}>
                          <h3 className={styles.fileName}>{row.name}</h3>
                          {row.isFolder ? (
                            <span className={styles.fileBadge}>Folder</span>
                          ) : null}
                        </div>
                        <ul className={styles.fileMeta}>
                          <li>
                            Candidate:{" "}
                            <strong>{row.candidate ?? "—"}</strong>
                          </li>
                          <li>
                            Visible:{" "}
                            <strong>
                              {row.customerVisible ? "Yes" : "No"}
                            </strong>
                          </li>
                          <li>
                            Modified:{" "}
                            <strong>
                              {row.modifiedDate
                                ? formatDisplayDate(row.modifiedDate)
                                : "—"}
                            </strong>
                          </li>
                          <li>
                            Status:{" "}
                            <strong
                              className={metadataStatusClass(row.metadataStatus)}
                            >
                              {row.metadataStatus}
                            </strong>
                          </li>
                        </ul>
                      </div>
                      <DocumentActionsMenu
                        row={row}
                        busy={busyId === row.id}
                        onEditMetadata={() => onEditMetadata(row)}
                        onSetVisibility={(customerVisible) =>
                          onSetVisibility(row, customerVisible)
                        }
                      />
                    </motion.article>
                  ))}
                </div>
              )
            ) : null}
          </motion.div>
        </AnimatePresence>
      </div>
    </div>
  );
}

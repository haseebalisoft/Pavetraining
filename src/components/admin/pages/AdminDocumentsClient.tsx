"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { motion } from "framer-motion";

import { AdminDrawer } from "@/components/admin/AdminDrawer";
import { DocumentsTableView } from "@/components/admin/documents/DocumentsTableView";
import styles from "@/components/admin/admin.module.css";
import docStyles from "@/components/admin/documents/documentsBrowse.module.css";
import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import {
  candidateDocumentsFolderName,
  companyDocumentsFolderName,
  resolveDocumentTypeFolder,
} from "@/lib/services/documentFolderPaths";
import {
  CUSTOMER_DOCUMENT_TYPES,
} from "@/types/adminDocuments";
import type { AdminDocumentRecord } from "@/types/adminDocuments";
import type { Company } from "@/types/models";

type WorkforceOption = {
  id: string;
  candidateName: string;
  companyName: string;
  workforceNumber: string | null;
};

type FormState = {
  companyId: string;
  candidateId: string;
  documentType: string;
  customerVisible: boolean;
  notificationSent: boolean;
  notifyCustomer: boolean;
};

function buildForm(row: AdminDocumentRecord): FormState {
  return {
    companyId: row.companyId ?? "",
    candidateId: row.candidateId ?? "",
    documentType: row.documentType ?? "",
    customerVisible: Boolean(row.customerVisible),
    notificationSent: Boolean(row.notificationSent),
    notifyCustomer: Boolean(row.notifyCustomer),
  };
}

/** Reverse of resolveDocumentTypeFolder: candidate subfolder → Document Type. */
const SUBFOLDER_TO_DOCUMENT_TYPE: Record<string, string> = {
  Certificates: "Certificate",
  "Card Scans": "Card Scan",
  "NVQ Documents": "NVQ Document",
  "Other Documents": "Other",
};

/** "C00002 - ali" → "C00002" (the stable number prefix before " - "). */
function folderNumberPrefix(segment: string | undefined | null): string {
  if (!segment) return "";
  const idx = segment.indexOf(" - ");
  return (idx >= 0 ? segment.slice(0, idx) : segment).trim();
}

/**
 * Admin Customer Documents — SharePoint folder tree browse:
 * Company → Company Documents | Candidates → Candidate → type folders.
 */
export function AdminDocumentsClient({
  companies,
  initialRows,
  initialPath = [],
  initialWorkforce,
}: {
  companies: Company[];
  initialRows: AdminDocumentRecord[];
  initialPath?: string[];
  initialWorkforce: WorkforceOption[];
}) {
  const { pushToast } = useAdminToast();
  const [rows, setRows] = useState(initialRows);
  const [workforce, setWorkforce] = useState(initialWorkforce);
  const [loading, setLoading] = useState(false);
  const [folderPath, setFolderPath] = useState<string[]>(initialPath);
  const [search, setSearch] = useState("");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [busyId, setBusyId] = useState<string | null>(null);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editing, setEditing] = useState<AdminDocumentRecord | null>(null);
  const [form, setForm] = useState<FormState>({
    companyId: "",
    candidateId: "",
    documentType: "",
    customerVisible: true,
    notificationSent: false,
    notifyCustomer: false,
  });
  const [formError, setFormError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [bulkOpen, setBulkOpen] = useState(false);
  const [bulkForm, setBulkForm] = useState<FormState>({
    companyId: "",
    candidateId: "",
    documentType: "",
    customerVisible: true,
    notificationSent: false,
    notifyCustomer: false,
  });
  const [bulkSaving, setBulkSaving] = useState(false);
  const [uploadOpen, setUploadOpen] = useState(false);
  const [uploadCompanyId, setUploadCompanyId] = useState("");
  const [uploadCandidateId, setUploadCandidateId] = useState("");
  const [uploadDocumentType, setUploadDocumentType] = useState<string>(
    "Certificate",
  );
  const [uploadCustomerVisible, setUploadCustomerVisible] = useState(true);
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [uploading, setUploading] = useState(false);

  const loadFolder = useCallback(
    async (path: string[]) => {
      setLoading(true);
      try {
        const params = new URLSearchParams({
          path: JSON.stringify(path),
        });
        const [docsResponse, workforceResponse] = await Promise.all([
          fetch(`/api/admin/documents/browse?${params}`, {
            cache: "no-store",
          }),
          fetch("/api/admin/workforce", { cache: "no-store" }),
        ]);
        if (!docsResponse.ok) {
          throw new Error(await readPublicApiError(docsResponse));
        }
        const docsPayload = (await docsResponse.json()) as {
          records?: AdminDocumentRecord[];
        };
        setRows(docsPayload.records ?? []);
        setFolderPath(path);
        setSelectedIds(new Set());
        if (workforceResponse.ok) {
          const workforcePayload = (await workforceResponse.json()) as {
            records?: WorkforceOption[];
          };
          setWorkforce(workforcePayload.records ?? []);
        }
      } catch (error) {
        pushToast(
          error instanceof Error ? error.message : "Failed to open folder",
          "error",
        );
      } finally {
        setLoading(false);
      }
    },
    [pushToast],
  );

  const load = useCallback(async () => {
    await loadFolder(folderPath);
  }, [folderPath, loadFolder]);

  const currentRows = useMemo(() => {
    const query = search.trim().toLowerCase();
    return rows
      .filter((row) => {
        if (!query) return true;
        return [
          row.name,
          row.company,
          row.candidate,
          row.documentType,
          row.id,
          row.modifiedBy,
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort((a, b) => {
        if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
        return a.name.localeCompare(b.name);
      });
  }, [rows, search]);

  const folderHint = useMemo(() => {
    if (folderPath.length === 0) {
      return "Root: company folders (Company Number - Company Name).";
    }
    if (folderPath.length === 1) {
      return "Company folder: open Company Documents or Candidates.";
    }
    if (folderPath.length === 2 && /candidates/i.test(folderPath[1] ?? "")) {
      return "Candidates: open a Candidate Number - Candidate Name folder.";
    }
    if (
      folderPath.length === 3 &&
      /candidates/i.test(folderPath[1] ?? "")
    ) {
      return "Candidate folder: Certificates, Card Scans, NVQ Documents, Other Documents.";
    }
    return "Browse files in this folder. Assign Company / Candidate metadata on files.";
  }, [folderPath]);

  const companyById = useMemo(() => {
    const map = new Map<string, Company>();
    for (const company of companies) map.set(company.id, company);
    return map;
  }, [companies]);

  const candidatesForCompany = useCallback(
    (companyId: string) => {
      const company = companyById.get(companyId);
      if (!company) return [] as WorkforceOption[];
      const name = company.companyName.trim().toLowerCase();
      return workforce
        .filter((row) => row.companyName.trim().toLowerCase() === name)
        .sort((a, b) => a.candidateName.localeCompare(b.candidateName));
    },
    [companyById, workforce],
  );

  const editCandidates = useMemo(
    () => candidatesForCompany(form.companyId),
    [candidatesForCompany, form.companyId],
  );
  const bulkCandidates = useMemo(
    () => candidatesForCompany(bulkForm.companyId),
    [candidatesForCompany, bulkForm.companyId],
  );
  const uploadCandidates = useMemo(
    () => candidatesForCompany(uploadCompanyId),
    [candidatesForCompany, uploadCompanyId],
  );

  // Seed the upload dialog from the folder currently open so a file lands in
  // THIS folder. Previously the dialog defaulted to companies[0] + no candidate,
  // so uploading inside a candidate subfolder (e.g. Card Scans) silently went to
  // the company's "Company Documents" folder and never appeared where expected.
  const deriveUploadSeed = useCallback(() => {
    const seed = {
      companyId: companies[0]?.id ?? "",
      candidateId: "",
      documentType: "Certificate",
    };
    const companySegment = folderPath[0];
    if (!companySegment) return seed;

    const companyMatch =
      companies.find(
        (c) =>
          companyDocumentsFolderName(c.companyNumber, c.companyName) ===
          companySegment,
      ) ??
      companies.find(
        (c) =>
          (c.companyNumber ?? "").trim() === folderNumberPrefix(companySegment),
      );
    if (!companyMatch) return seed;
    seed.companyId = companyMatch.id;

    if (/^candidates$/i.test((folderPath[1] ?? "").trim())) {
      const candidateSegment = folderPath[2];
      if (candidateSegment) {
        const companyKey = companyMatch.companyName.trim().toLowerCase();
        const candidateMatch =
          workforce.find(
            (w) =>
              w.companyName.trim().toLowerCase() === companyKey &&
              candidateDocumentsFolderName(
                w.workforceNumber,
                w.candidateName,
              ) === candidateSegment,
          ) ??
          workforce.find(
            (w) =>
              w.companyName.trim().toLowerCase() === companyKey &&
              (w.workforceNumber ?? "").trim() ===
                folderNumberPrefix(candidateSegment),
          );
        if (candidateMatch) seed.candidateId = candidateMatch.id;
      }
      const subfolder = folderPath[3];
      if (subfolder && SUBFOLDER_TO_DOCUMENT_TYPE[subfolder]) {
        seed.documentType = SUBFOLDER_TO_DOCUMENT_TYPE[subfolder];
      }
    }
    return seed;
  }, [companies, workforce, folderPath]);

  // Live preview of the exact SharePoint folder the file will be stored in,
  // built from the current selections with the same rules as the server.
  const uploadDestination = useMemo(() => {
    const company = companyById.get(uploadCompanyId);
    if (!company) return null;
    const hasCandidate = Boolean(uploadCandidateId);
    const destinationFolder = resolveDocumentTypeFolder({
      documentType: uploadDocumentType,
      hasCandidate,
    });
    const companyFolder = companyDocumentsFolderName(
      company.companyNumber,
      company.companyName,
    );
    if (!hasCandidate || destinationFolder === "Company Documents") {
      return `${companyFolder}/Company Documents`;
    }
    const candidate = workforce.find((w) => w.id === uploadCandidateId);
    const candidateFolder = candidate
      ? candidateDocumentsFolderName(
          candidate.workforceNumber,
          candidate.candidateName,
        )
      : "Candidate";
    return `${companyFolder}/Candidates/${candidateFolder}/${destinationFolder}`;
  }, [
    companyById,
    uploadCompanyId,
    uploadCandidateId,
    uploadDocumentType,
    workforce,
  ]);

  useEffect(() => {
    if (
      form.candidateId &&
      !editCandidates.some((row) => row.id === form.candidateId)
    ) {
      setForm((current) => ({ ...current, candidateId: "" }));
    }
  }, [editCandidates, form.candidateId]);

  function openFolder(row: AdminDocumentRecord) {
    if (!row.isFolder) return;
    void loadFolder([...folderPath, row.name]);
  }

  function navigateToCrumb(index: number) {
    if (index < 0) {
      void loadFolder([]);
      return;
    }
    void loadFolder(folderPath.slice(0, index + 1));
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
          ? "Customer Visible set to Yes."
          : "Customer Visible set to No.",
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

  async function persistEdit() {
    if (!editing) return;
    setSaving(true);
    setFormError(null);
    try {
      const response = await fetch(`/api/admin/documents/${editing.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          companyId: form.companyId || null,
          candidateId: form.candidateId || null,
          documentType: form.documentType || null,
          customerVisible: form.customerVisible,
          notificationSent: form.notificationSent,
          notifyCustomer: form.notifyCustomer,
        }),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      pushToast("Document metadata saved.");
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

  async function deleteDocument() {
    if (!editing || editing.isFolder) return;
    if (
      !window.confirm(
        `Delete “${editing.name}”?\n\nThis removes the file from SharePoint Customer Documents.`,
      )
    ) {
      return;
    }
    setSaving(true);
    setFormError(null);
    try {
      const response = await fetch(`/api/admin/documents/${editing.id}`, {
        method: "DELETE",
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      pushToast("Document deleted.");
      setDrawerOpen(false);
      setEditing(null);
      await load();
    } catch (error) {
      const message = error instanceof Error ? error.message : "Delete failed.";
      setFormError(message);
      pushToast(message, "error");
    } finally {
      setSaving(false);
    }
  }

  async function persistBulk() {
    const ids = Array.from(selectedIds);
    if (ids.length === 0) {
      pushToast("Select one or more files first.", "error");
      return;
    }
    setBulkSaving(true);
    try {
      const response = await fetch("/api/admin/documents/bulk", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ids,
          companyId: bulkForm.companyId || null,
          candidateId: bulkForm.candidateId || null,
          documentType: bulkForm.documentType || null,
          customerVisible: bulkForm.customerVisible,
          notificationSent: bulkForm.notificationSent,
          notifyCustomer: bulkForm.notifyCustomer,
        }),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      const payload = (await response.json()) as {
        updated?: unknown[];
        failed?: string[];
      };
      const failed = payload.failed?.length ?? 0;
      pushToast(
        failed
          ? `Updated ${payload.updated?.length ?? 0}; ${failed} failed.`
          : `Updated ${payload.updated?.length ?? ids.length} document(s).`,
        failed ? "error" : "success",
      );
      setBulkOpen(false);
      await load();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Bulk update failed.",
        "error",
      );
    } finally {
      setBulkSaving(false);
    }
  }

  async function persistUpload() {
    if (!uploadCompanyId) {
      setUploadError("Company is required.");
      return;
    }
    if (!uploadFile) {
      setUploadError("Choose a file to upload.");
      return;
    }
    if (uploadDocumentType !== "Other" && !uploadCandidateId) {
      // Company-level "Other" without candidate still allowed; Certificate etc need candidate when in candidate folders
    }

    setUploading(true);
    setUploadError(null);
    try {
      const body = new FormData();
      body.set("file", uploadFile);
      body.set("companyId", uploadCompanyId);
      body.set("documentType", uploadDocumentType);
      body.set("customerVisible", uploadCustomerVisible ? "true" : "false");
      if (uploadCandidateId) body.set("candidateId", uploadCandidateId);

      const response = await fetch("/api/admin/documents/upload", {
        method: "POST",
        body,
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      pushToast("Document uploaded.");
      setUploadOpen(false);
      await load();
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Upload failed.";
      setUploadError(message);
      pushToast(message, "error");
    } finally {
      setUploading(false);
    }
  }

  function renderAssignFields(
    state: FormState,
    setState: (next: FormState) => void,
    candidates: WorkforceOption[],
  ) {
    return (
      <div className={styles.formGrid}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Company</span>
          <select
            className={styles.select}
            value={state.companyId}
            onChange={(event) =>
              setState({
                ...state,
                companyId: event.target.value,
                candidateId: "",
              })
            }
          >
            <option value="">Select company…</option>
            {companies.map((company) => (
              <option key={company.id} value={company.id}>
                {company.companyName}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Candidate</span>
          <select
            className={styles.select}
            value={state.candidateId}
            disabled={!state.companyId}
            onChange={(event) =>
              setState({ ...state, candidateId: event.target.value })
            }
          >
            <option value="">No candidate (company-level)</option>
            {candidates.map((row) => (
              <option key={row.id} value={row.id}>
                {row.candidateName}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Document Type</span>
          <select
            className={styles.select}
            value={state.documentType}
            onChange={(event) =>
              setState({ ...state, documentType: event.target.value })
            }
          >
            <option value="">Select type…</option>
            {CUSTOMER_DOCUMENT_TYPES.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </label>
        <label className={styles.checkboxRow}>
          <input
            type="checkbox"
            checked={state.customerVisible}
            onChange={(event) =>
              setState({ ...state, customerVisible: event.target.checked })
            }
          />
          Customer Visible
        </label>
        <label className={styles.checkboxRow}>
          <input
            type="checkbox"
            checked={state.notificationSent}
            onChange={(event) =>
              setState({ ...state, notificationSent: event.target.checked })
            }
          />
          Notification Sent
        </label>
        <label className={styles.checkboxRow}>
          <input
            type="checkbox"
            checked={state.notifyCustomer}
            onChange={(event) =>
              setState({ ...state, notifyCustomer: event.target.checked })
            }
          />
          Notify Customer
        </label>
      </div>
    );
  }

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={[
              { label: "Admin", href: "/admin" },
              { label: "Customer Documents" },
            ]}
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Customer Documents</h1>
          <p className={styles.subtitle}>
            Same SharePoint folder tree: Company Number - Company Name → Company
            Documents / Candidates → Candidate Number - Name → Certificates,
            Card Scans, NVQ Documents, Other Documents.
          </p>
          <p className={styles.helpText}>{folderHint}</p>
        </div>
        <div className={styles.headerActions}>
          <motion.button
            type="button"
            className={styles.secondaryButton}
            disabled={selectedIds.size === 0}
            whileTap={{ scale: 0.97, transition: { duration: 0.15 } }}
            onClick={() => {
              setBulkForm({
                companyId: "",
                candidateId: "",
                documentType: "",
                customerVisible: true,
                notificationSent: false,
                notifyCustomer: false,
              });
              setBulkOpen(true);
            }}
          >
            Assign selected ({selectedIds.size})
          </motion.button>
          <motion.button
            type="button"
            className={styles.primaryButton}
            whileTap={{ scale: 0.97, transition: { duration: 0.15 } }}
            onClick={() => {
              const seed = deriveUploadSeed();
              setUploadOpen(true);
              setUploadError(null);
              setUploadFile(null);
              setUploadCompanyId(seed.companyId);
              setUploadCandidateId(seed.candidateId);
              setUploadDocumentType(seed.documentType);
              setUploadCustomerVisible(true);
            }}
          >
            Create or upload
          </motion.button>
        </div>
      </header>

      <nav className={docStyles.browseCrumbs} aria-label="Folder location">
        <button
          type="button"
          className={
            folderPath.length === 0
              ? docStyles.crumbCurrent
              : docStyles.crumbButton
          }
          disabled={folderPath.length === 0}
          onClick={() => navigateToCrumb(-1)}
        >
          Customer Documents
        </button>
        {folderPath.map((segment, index) => (
          <span key={`${segment}-${index}`} className={docStyles.crumbSeg}>
            <span className={docStyles.crumbSep} aria-hidden="true">
              /
            </span>
            <button
              type="button"
              className={
                index === folderPath.length - 1
                  ? docStyles.crumbCurrent
                  : docStyles.crumbButton
              }
              disabled={index === folderPath.length - 1}
              onClick={() => navigateToCrumb(index)}
            >
              {segment}
            </button>
          </span>
        ))}
      </nav>

      <div className={styles.crudToolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            value={search}
            placeholder="Filter this folder…"
            onChange={(event) => setSearch(event.target.value)}
          />
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

      <p className={styles.resultMeta}>
        {currentRows.length} item{currentRows.length === 1 ? "" : "s"}
        {selectedIds.size > 0 ? ` · ${selectedIds.size} selected` : ""}
      </p>

      {loading ? (
        <LoadingState label="Refreshing Customer Documents…" />
      ) : (
        <DocumentsTableView
          rows={currentRows}
          selectedIds={selectedIds}
          busyId={busyId}
          onToggleSelect={(id) => {
            setSelectedIds((current) => {
              const next = new Set(current);
              if (next.has(id)) next.delete(id);
              else next.add(id);
              return next;
            });
          }}
          onToggleSelectAll={(selectAll) => {
            if (!selectAll) {
              setSelectedIds(new Set());
              return;
            }
            setSelectedIds(
              new Set(
                currentRows.filter((row) => !row.isFolder).map((row) => row.id),
              ),
            );
          }}
          onOpenFolder={openFolder}
          onEditMetadata={openEdit}
          onSetVisibility={setVisibility}
        />
      )}

      <AdminDrawer
        open={drawerOpen}
        title="Edit document fields"
        onClose={() => setDrawerOpen(false)}
        footer={
          <>
            {editing && !editing.isFolder ? (
              <button
                type="button"
                className={styles.dangerButton}
                disabled={saving}
                onClick={() => void deleteDocument()}
              >
                {saving ? "Working…" : "Delete"}
              </button>
            ) : null}
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
              onClick={() => void persistEdit()}
            >
              {saving ? "Saving…" : "Save"}
            </motion.button>
          </>
        }
      >
        {editing ? (
          <p className={styles.helpText}>
            {editing.name} · ID {editing.id}
          </p>
        ) : null}
        {formError ? <p className={styles.formError}>{formError}</p> : null}
        {renderAssignFields(form, setForm, editCandidates)}
      </AdminDrawer>

      <AdminDrawer
        open={bulkOpen}
        title={`Assign ${selectedIds.size} document(s)`}
        onClose={() => setBulkOpen(false)}
        footer={
          <>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setBulkOpen(false)}
            >
              Cancel
            </button>
            <motion.button
              type="button"
              className={styles.primaryButton}
              disabled={bulkSaving}
              whileTap={{ scale: 0.97, transition: { duration: 0.15 } }}
              onClick={() => void persistBulk()}
            >
              {bulkSaving ? "Saving…" : "Save"}
            </motion.button>
          </>
        }
      >
        <p className={styles.helpText}>
          Set Company, Candidate, Document Type, and Customer Visible on all
          selected files (same SharePoint fields).
        </p>
        {renderAssignFields(bulkForm, setBulkForm, bulkCandidates)}
      </AdminDrawer>

      <AdminDrawer
        open={uploadOpen}
        title="Create or upload"
        onClose={() => setUploadOpen(false)}
        footer={
          <>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setUploadOpen(false)}
            >
              Cancel
            </button>
            <motion.button
              type="button"
              className={styles.primaryButton}
              disabled={uploading}
              whileTap={{ scale: 0.97, transition: { duration: 0.15 } }}
              onClick={() => void persistUpload()}
            >
              {uploading ? "Uploading…" : "Upload"}
            </motion.button>
          </>
        }
      >
        {uploadError ? <p className={styles.formError}>{uploadError}</p> : null}
        <div className={styles.formGrid}>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Company</span>
            <select
              className={styles.select}
              value={uploadCompanyId}
              onChange={(event) => {
                setUploadCompanyId(event.target.value);
                setUploadCandidateId("");
              }}
            >
              <option value="">Select company…</option>
              {companies.map((company) => (
                <option key={company.id} value={company.id}>
                  {company.companyName}
                </option>
              ))}
            </select>
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Candidate</span>
            <select
              className={styles.select}
              value={uploadCandidateId}
              disabled={!uploadCompanyId}
              onChange={(event) => setUploadCandidateId(event.target.value)}
            >
              <option value="">Company-level document</option>
              {uploadCandidates.map((row) => (
                <option key={row.id} value={row.id}>
                  {row.candidateName}
                </option>
              ))}
            </select>
          </label>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Document Type</span>
            <select
              className={styles.select}
              value={uploadDocumentType}
              onChange={(event) => setUploadDocumentType(event.target.value)}
            >
              {CUSTOMER_DOCUMENT_TYPES.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </label>
          {uploadDestination ? (
            <p className={styles.helpText}>
              Uploads to: <strong>{uploadDestination}</strong>
            </p>
          ) : null}
          <label className={styles.field}>
            <span className={styles.fieldLabel}>File</span>
            <input
              className={styles.input}
              type="file"
              onChange={(event) =>
                setUploadFile(event.target.files?.[0] ?? null)
              }
            />
          </label>
          <label className={styles.checkboxRow}>
            <input
              type="checkbox"
              checked={uploadCustomerVisible}
              onChange={(event) =>
                setUploadCustomerVisible(event.target.checked)
              }
            />
            Customer Visible
          </label>
        </div>
      </AdminDrawer>
    </div>
  );
}

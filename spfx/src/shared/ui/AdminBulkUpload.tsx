import * as React from "react";

import type { SharePointListKey } from "../schema/sharepointSchema";
import { getSharePointFields } from "../schema/sharepointSchema";
import {
  createListItem,
  getListItems,
  updateListItem,
  type SpListClient,
} from "../services/sharePointListService";
import styles from "./portal.module.scss";

type ImportType =
  | "company"
  | "workforce"
  | "trainingMatrix"
  | "npors"
  | "eusr"
  | "streetworks"
  | "inHouse"
  | "nvq";

type DuplicateMode = "skip" | "update" | "create";

interface PreviewRow {
  rowNumber: number;
  status: "Ready" | "Warning" | "Duplicate" | "Error";
  message: string;
  fields: Record<string, string>;
  existingId?: string;
}

const IMPORT_OPTIONS: Array<{
  value: ImportType;
  label: string;
  hint: string;
  listKey: SharePointListKey;
  matchField: string;
  templateHeaders: string[];
}> = [
  {
    value: "company",
    label: "Companies",
    hint: "CSV with Company Name, Company Number, Email, Status…",
    listKey: "company",
    matchField: "CompanyName",
    templateHeaders: [
      "Company Name",
      "Company Number",
      "Company Size",
      "Email",
      "Tel No",
      "Main Contact",
      "Status",
    ],
  },
  {
    value: "workforce",
    label: "Workforce / Candidates",
    hint: "CSV with Candidate Name, Company Name, Email, Status…",
    listKey: "workforce",
    matchField: "CandidateName",
    templateHeaders: [
      "Candidate Name",
      "Company Name",
      "Workforce Number",
      "Email",
      "Department",
      "Status",
    ],
  },
  {
    value: "trainingMatrix",
    label: "Training Matrix Update",
    hint: "CSV with Name (Title), DOB, CSCS Expiry, and category date columns.",
    listKey: "trainingMatrix",
    matchField: "Title",
    templateHeaders: [
      "Name",
      "DOB",
      "CSCS Expiry",
      "SSSTS Expiry",
      "SMSTS Expiry",
      "NRSWA Expiry",
      "EUSR Expiry",
    ],
  },
  {
    value: "npors",
    label: "NPORS records",
    hint: "CSV with Candidate Name, Company Name, NPORS Number, Outcome…",
    listKey: "nporsRegister",
    matchField: "CandidateName",
    templateHeaders: [
      "Candidate Name",
      "Company Name",
      "NPORS Number",
      "NPORS Category",
      "Training Outcome",
      "Expiry",
    ],
  },
  {
    value: "eusr",
    label: "EUSR records",
    hint: "CSV with Candidate Name, Company Name, EUSR Number, Outcome…",
    listKey: "eusrRegister",
    matchField: "CandidateName",
    templateHeaders: [
      "Candidate Name",
      "Company Name",
      "EUSR Number",
      "Eusr Category",
      "Training Outcome",
      "Expiry",
    ],
  },
  {
    value: "streetworks",
    label: "Streetworks / NRSWA",
    hint: "CSV with Candidate Name, Company Name, SWQR Number, Course…",
    listKey: "nrswaRegister",
    matchField: "CandidateName",
    templateHeaders: [
      "Candidate Name",
      "Company Name",
      "SWQR Number",
      "Course",
      "Training Outcome",
      "Expiry date",
    ],
  },
  {
    value: "inHouse",
    label: "In-House records",
    hint: "CSV with Candidate Name, Company Name, Certificate Category…",
    listKey: "inHouseCertificates",
    matchField: "CandidateName",
    templateHeaders: [
      "Candidate Name",
      "Company Name",
      "Certificate Category",
      "Course Category",
      "Training Outcome",
      "Expiry Date",
    ],
  },
  {
    value: "nvq",
    label: "NVQ records",
    hint: "CSV with Candidate Name, NVQ Title, Stage…",
    listKey: "nvqRegister",
    matchField: "CandidateName",
    templateHeaders: [
      "Candidate Name",
      "Company Name",
      "Nvq Title",
      "Stage of Nvq",
      "Customer Visible",
    ],
  },
];

/** Map common spreadsheet headers → SharePoint internal names. */
const HEADER_ALIASES: Record<string, Record<string, string>> = {
  company: {
    "company name": "CompanyName",
    "company number": "CompanyNumber",
    "company size": "CompanySize",
    email: "Email",
    "tel no": "TelNo",
    "main contact": "MainContact",
    status: "Status",
    title: "Title",
  },
  workforce: {
    "candidate name": "CandidateName",
    "company name": "CompanyName",
    "workforce number": "WorkforceNumber",
    email: "Email",
    department: "Department",
    status: "Status",
    "date of birth": "DateofBirth",
    "cscs number": "CSCSNumber",
    "eusr number": "EUSRNumber",
    "npors numbers": "NPORSNumbers",
    "swqr number": "SWQRNumber",
  },
  trainingMatrix: {
    name: "Title",
    title: "Title",
    dob: "DOB",
    "cscs expiry": "CSCSExpiry",
    "sssts expiry": "SSSTSExpiry",
    "smsts expiry": "SMSTSExpiry",
    "nrswa expiry": "NRSWAExpiry",
    "eusr expiry": "EUSRExpiry",
    "face ift": "Faceift",
  },
  npors: {
    "candidate name": "CandidateName",
    "company name": "CompanyName",
    "npors number": "NPORSNumber",
    "npors category": "NPORSCategory",
    "training outcome": "TrainingOutcome",
    expiry: "Expiry",
    "training date": "TrainingDate",
  },
  eusr: {
    "candidate name": "CandidateName",
    "company name": "CompanyName",
    "eusr number": "EUSRNumber",
    "eusr category": "EusrCategory",
    "training outcome": "TrainingOutcome",
    expiry: "Expiry",
  },
  streetworks: {
    "candidate name": "CandidateName",
    "company name": "CompanyName",
    "swqr number": "SWQRNumber",
    course: "Course",
    "training outcome": "TrainingOutcome",
    "expiry date": "Expirydate",
  },
  inHouse: {
    "candidate name": "CandidateName",
    "company name": "CompanyName",
    "certificate category": "CertificateCategory",
    "course category": "CourseCategory",
    "training outcome": "TrainingOutcome",
    "expiry date": "ExpiryDate",
  },
  nvq: {
    "candidate name": "CandidateName",
    "company name": "Company_x0020_Name",
    "nvq title": "NvqTitle",
    "stage of nvq": "StageofNvq",
    "customer visible": "CustomerVisible",
  },
};

function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;
  const src = text.replace(/^\uFEFF/, "");
  for (let i = 0; i < src.length; i++) {
    const ch = src.charAt(i);
    const next = src.charAt(i + 1);
    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (ch === "\r") {
      // skip
    } else {
      cell += ch;
    }
  }
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows.filter((r) => r.some((c) => c.trim() !== ""));
}

function downloadTemplate(opt: (typeof IMPORT_OPTIONS)[number]): void {
  const csv = opt.templateHeaders.join(",") + "\n";
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = opt.value + "-template.csv";
  a.click();
  URL.revokeObjectURL(url);
}

function toIsoDate(raw: string): string | undefined {
  const t = raw.trim();
  if (!t) return undefined;
  if (/^\d{4}-\d{2}-\d{2}/.test(t)) return t.slice(0, 10) + "T00:00:00Z";
  const m = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    const iso =
      y +
      "-" +
      (mo < 10 ? "0" : "") +
      String(mo) +
      "-" +
      (d < 10 ? "0" : "") +
      String(d);
    return iso + "T00:00:00Z";
  }
  const parsed = new Date(t);
  if (!isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10) + "T00:00:00Z";
  }
  return undefined;
}

export interface AdminBulkUploadProps {
  client: SpListClient;
}

export const AdminBulkUpload: React.FC<AdminBulkUploadProps> = (props) => {
  const { client } = props;
  const [importType, setImportType] = React.useState<ImportType>("company");
  const [duplicateMode, setDuplicateMode] =
    React.useState<DuplicateMode>("skip");
  const [preview, setPreview] = React.useState<PreviewRow[]>([]);
  const [fileName, setFileName] = React.useState<string>("");
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [ok, setOk] = React.useState<string | null>(null);

  const option =
    IMPORT_OPTIONS.filter((o) => o.value === importType)[0] ||
    IMPORT_OPTIONS[0];

  const summary = React.useMemo(() => {
    const s = {
      total: preview.length,
      ready: 0,
      warning: 0,
      duplicate: 0,
      error: 0,
    };
    for (let i = 0; i < preview.length; i++) {
      const st = preview[i].status;
      if (st === "Ready") s.ready++;
      else if (st === "Warning") s.warning++;
      else if (st === "Duplicate") s.duplicate++;
      else s.error++;
    }
    return s;
  }, [preview]);

  const buildPreview = async (file: File): Promise<void> => {
    setBusy(true);
    setError(null);
    setOk(null);
    setFileName(file.name);
    try {
      const text = await file.text();
      const table = parseCsv(text);
      if (table.length < 2) {
        throw new Error("CSV needs a header row and at least one data row.");
      }
      const headers = table[0].map((h) => h.trim());
      const aliases = HEADER_ALIASES[importType] || {};
      const mappedHeaders = headers.map((h) => {
        const key = h.toLowerCase();
        return aliases[key] || "";
      });

      const existing = await getListItems(client, option.listKey, {
        top: 5000,
        maxItems: 20000,
      });
      const byMatch: Record<string, string> = {};
      for (let i = 0; i < existing.length; i++) {
        const v = String(
          existing[i].fields[option.matchField] ||
            existing[i].fields.Title ||
            ""
        )
          .trim()
          .toLowerCase();
        if (v) byMatch[v] = existing[i].id;
      }

      const rows: PreviewRow[] = [];
      for (let r = 1; r < table.length; r++) {
        const cells = table[r];
        const fields: Record<string, string> = {};
        for (let c = 0; c < mappedHeaders.length; c++) {
          const internal = mappedHeaders[c];
          if (!internal) continue;
          const raw = (cells[c] || "").trim();
          if (!raw) continue;
          if (/expiry|dob|date/i.test(internal)) {
            const iso = toIsoDate(raw);
            fields[internal] = iso || raw;
          } else if (internal === "CustomerVisible") {
            fields[internal] =
              /^(1|true|yes|y)$/i.test(raw) ? "true" : "false";
          } else {
            fields[internal] = raw;
          }
        }

        const matchVal = String(
          fields[option.matchField] || fields.Title || ""
        ).trim();
        if (!matchVal) {
          rows.push({
            rowNumber: r + 1,
            status: "Error",
            message: "Missing required match field (" + option.matchField + ")",
            fields,
          });
          continue;
        }

        const existingId = byMatch[matchVal.toLowerCase()];
        if (existingId) {
          rows.push({
            rowNumber: r + 1,
            status: "Duplicate",
            message: "Matches existing item #" + existingId,
            fields,
            existingId,
          });
        } else {
          rows.push({
            rowNumber: r + 1,
            status: "Ready",
            message: "Ready to create",
            fields,
          });
        }
      }
      setPreview(rows);
    } catch (e) {
      setPreview([]);
      setError(e instanceof Error ? e.message : "Failed to preview file.");
    } finally {
      setBusy(false);
    }
  };

  const commit = async (): Promise<void> => {
    setBusy(true);
    setError(null);
    setOk(null);
    try {
      let created = 0;
      let updated = 0;
      let skipped = 0;
      let failed = 0;

      for (let i = 0; i < preview.length; i++) {
        const row = preview[i];
        if (row.status === "Error") {
          failed++;
          continue;
        }

        const payload: Record<string, unknown> = {};
        for (const key in row.fields) {
          if (!Object.prototype.hasOwnProperty.call(row.fields, key)) continue;
          const val = row.fields[key];
          if (key === "CustomerVisible") {
            payload[key] = val === "true";
          } else {
            payload[key] = val;
          }
        }

        // Ensure Title for company / matrix creates
        if (
          option.listKey === "company" &&
          !payload.Title &&
          payload.CompanyName
        ) {
          payload.Title = payload.CompanyName;
        }

        try {
          if (row.existingId) {
            if (duplicateMode === "skip") {
              skipped++;
              continue;
            }
            if (duplicateMode === "update") {
              await updateListItem(
                client,
                option.listKey,
                row.existingId,
                payload
              );
              updated++;
              continue;
            }
          }
          await createListItem(client, option.listKey, payload);
          created++;
        } catch {
          failed++;
        }
      }

      setOk(
        "Import finished — created " +
          created +
          ", updated " +
          updated +
          ", skipped " +
          skipped +
          ", failed " +
          failed +
          "."
      );
      setPreview([]);
      setFileName("");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Commit failed.");
    } finally {
      setBusy(false);
    }
  };

  // Touch schema so list field maps stay discoverable for future importers.
  void getSharePointFields("company");

  return (
    <div className={styles.panel}>
      <p className={styles.muted}>
        Same flow as Next.js admin: choose type → download template → preview
        CSV → commit. Save Excel files as CSV before upload. Writes directly to
        SharePoint lists (including Training Matrix Update).
      </p>

      <div className={styles.bulkGrid}>
        {IMPORT_OPTIONS.map((opt) => (
          <button
            key={opt.value}
            type="button"
            className={
              importType === opt.value
                ? styles.bulkTypeActive
                : styles.bulkTypeBtn
            }
            onClick={() => {
              setImportType(opt.value);
              setPreview([]);
              setOk(null);
              setError(null);
            }}
          >
            <strong>{opt.label}</strong>
            <span>{opt.hint}</span>
          </button>
        ))}
      </div>

      <div className={styles.toolbar}>
        <button type="button" onClick={() => downloadTemplate(option)}>
          Download CSV template
        </button>
        <label className={styles.bulkFileLabel}>
          Choose CSV
          <input
            type="file"
            accept=".csv,text/csv"
            disabled={busy}
            onChange={(e) => {
              const f = e.target.files && e.target.files[0];
              if (f) {
                buildPreview(f).catch(() => undefined);
              }
              e.target.value = "";
            }}
          />
        </label>
        <select
          value={duplicateMode}
          onChange={(e) =>
            setDuplicateMode(e.target.value as DuplicateMode)
          }
          disabled={busy}
          aria-label="Duplicate handling"
        >
          <option value="skip">Duplicates: skip</option>
          <option value="update">Duplicates: update</option>
          <option value="create">Duplicates: create anyway</option>
        </select>
        <button
          type="button"
          onClick={() => {
            commit().catch(() => undefined);
          }}
          disabled={busy || preview.length === 0}
        >
          {busy ? "Working…" : "Commit import"}
        </button>
      </div>

      {fileName && (
        <p className={styles.muted}>
          File: {fileName} · {summary.total} rows ({summary.ready} ready,{" "}
          {summary.duplicate} duplicates, {summary.error} errors)
        </p>
      )}
      {error && <p className={styles.error}>{error}</p>}
      {ok && <p className={styles.success}>{ok}</p>}

      {preview.length > 0 && (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Row</th>
                <th>Status</th>
                <th>Message</th>
                <th>Key</th>
              </tr>
            </thead>
            <tbody>
              {preview.slice(0, 200).map((row) => (
                <tr key={row.rowNumber}>
                  <td>{row.rowNumber}</td>
                  <td>{row.status}</td>
                  <td title={row.message}>{row.message}</td>
                  <td
                    title={
                      row.fields[option.matchField] ||
                      row.fields.Title ||
                      ""
                    }
                  >
                    {row.fields[option.matchField] ||
                      row.fields.Title ||
                      "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

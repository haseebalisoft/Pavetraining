import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";

import { getSharePointFields } from "../schema/sharepointSchema";
import { getActivePermissionByEmail } from "../services/permissionService";
import {
  loadDashboardCounts,
  loadPortalListRows,
  type PortalTableRow,
} from "../services/portalDataService";
import type { PermissionProfile } from "../types/models";
import {
  ADMIN_NAV,
  CUSTOMER_NAV,
  type AdminViewId,
  type CustomerViewId,
} from "./nav";
import styles from "./portal.module.scss";

export type PortalMode = "admin" | "customer";

export interface PortalShellProps {
  mode: PortalMode;
  spHttpClient: SPHttpClient;
  webUrl: string;
  userEmail: string;
}

interface ViewSpec {
  title: string;
  subtitle: string;
  listKey?:
    | "company"
    | "workforce"
    | "trainingMatrix"
    | "nporsRegister"
    | "eusrRegister"
    | "nrswaRegister"
    | "inHouseCertificates"
    | "nvqRegister"
    | "customerDocuments"
    | "events"
    | "offersPromotions"
    | "permissions"
    | "trainingManagerLogs";
  columns?: string[];
  headers?: string[];
  companyField?: string;
  visibleField?: string;
  customerScoped?: boolean;
  stub?: string;
}

const ADMIN_VIEWS: Record<AdminViewId, ViewSpec> = {
  dashboard: {
    title: "Dashboard",
    subtitle: "Operations overview from SharePoint lists",
  },
  companies: {
    title: "Companies",
    subtitle: "Company List",
    listKey: "company",
    headers: ["Company", "Number", "Email", "Status"],
    columns: ["CompanyName", "CompanyNumber", "Email", "Status"],
  },
  workforce: {
    title: "Workforce",
    subtitle: "Workforce List",
    listKey: "workforce",
    headers: ["Candidate", "Company", "Department", "Status"],
    columns: ["CandidateName", "CompanyName", "Department", "Status"],
  },
  "training-matrix": {
    title: "Training Matrix",
    subtitle: "Matrix rows and expiry dates",
    listKey: "trainingMatrix",
    headers: ["Candidate", "Company", "Status", "Next expiry"],
    columns: [
      "CandidateName",
      "Company_x0020_Name",
      "OverallStatus",
      "NextExpiryDate",
    ],
  },
  "training-records": {
    title: "Training Records",
    subtitle: "Choose a register",
  },
  nvq: {
    title: "NVQ",
    subtitle: "NVQ Register",
    listKey: "nvqRegister",
    headers: ["Candidate", "Title", "Stage", "Company"],
    columns: [
      "CandidateName",
      "NvqTitle",
      "StageofNvq",
      "Company_x0020_Name",
    ],
  },
  documents: {
    title: "Documents",
    subtitle: "Customer Documents library",
    listKey: "customerDocuments",
    headers: ["Name", "Company", "Type", "Visible"],
    columns: ["FileLeafRef", "Company", "DocumentType", "CustomerVisible"],
  },
  events: {
    title: "Events",
    subtitle: "Events calendar list",
    listKey: "events",
    headers: ["Title", "Company", "Start", "Visible"],
    columns: ["Title", "EventCompany", "EventDate", "Customer_x0020_Visible"],
  },
  offers: {
    title: "Offers",
    subtitle: "Offers / Promotions",
    listKey: "offersPromotions",
    headers: ["Title", "Category", "Status", "Visible"],
    columns: ["Title", "Category", "Status", "CustomerVisible"],
  },
  permissions: {
    title: "Permissions",
    subtitle: "Portal access control",
    listKey: "permissions",
    headers: ["Email", "Role", "Status", "Scope"],
    columns: ["UserEmail", "RoleType", "Status", "AccessScope"],
  },
  automation: {
    title: "Automation",
    subtitle: "Automation rules and sync controls",
    stub: "Automation rules and Outlook sync controls will connect here (same concept as Next.js /admin/automation).",
  },
  logs: {
    title: "Logs",
    subtitle: "Training Manager Logs",
    listKey: "trainingManagerLogs",
    headers: ["Title", "User", "List", "Timestamp"],
    columns: ["Title", "User_x0020_Email", "ListName", "Timestamp"],
  },
};

const CUSTOMER_VIEWS: Record<CustomerViewId, ViewSpec> = {
  dashboard: {
    title: "Dashboard",
    subtitle: "Your company training overview",
  },
  "training-matrix": {
    title: "Training Matrix",
    subtitle: "Your company's matrix rows",
    listKey: "trainingMatrix",
    headers: ["Candidate", "Status", "Next expiry"],
    columns: ["CandidateName", "OverallStatus", "NextExpiryDate"],
    companyField: "Company_x0020_Name",
    customerScoped: true,
  },
  candidates: {
    title: "Candidates",
    subtitle: "Workforce for your company",
    listKey: "workforce",
    headers: ["Candidate", "Department", "Status"],
    columns: ["CandidateName", "Department", "Status"],
    companyField: "CompanyName",
    customerScoped: true,
  },
  "training-records": {
    title: "Training Records",
    subtitle: "Choose a register",
  },
  "nvq-progress": {
    title: "NVQ Progress",
    subtitle: "Visible NVQ records for your company",
    listKey: "nvqRegister",
    headers: ["Candidate", "Title", "Stage"],
    columns: ["CandidateName", "NvqTitle", "StageofNvq"],
    companyField: "Company_x0020_Name",
    visibleField: "CustomerVisible",
    customerScoped: true,
  },
  documents: {
    title: "Documents",
    subtitle: "Customer-visible files",
    listKey: "customerDocuments",
    headers: ["Name", "Type", "Candidate"],
    columns: ["FileLeafRef", "DocumentType", "Candidate"],
    companyField: "Company",
    visibleField: "CustomerVisible",
    customerScoped: true,
  },
  events: {
    title: "Events",
    subtitle: "Visible upcoming events",
    listKey: "events",
    headers: ["Title", "Start", "Address"],
    columns: ["Title", "EventDate", "TrainingAddress"],
    companyField: "EventCompany",
    visibleField: "Customer_x0020_Visible",
    customerScoped: true,
  },
  offers: {
    title: "Offers",
    subtitle: "Visible offers and promotions",
    listKey: "offersPromotions",
    headers: ["Title", "Category", "Status"],
    columns: ["Title", "Category", "Status"],
    visibleField: "CustomerVisible",
    customerScoped: true,
  },
  support: {
    title: "Support",
    subtitle: "Contact your PAVE training manager",
    stub: "For renewals, certificates, or portal access, contact your assigned Training Manager. Include your company name and login email.",
  },
};

const RECORD_SUBNAV = [
  { id: "npors" as const, label: "NPORS", listKey: "nporsRegister" as const },
  { id: "eusr" as const, label: "EUSR", listKey: "eusrRegister" as const },
  {
    id: "streetworks" as const,
    label: "Streetworks",
    listKey: "nrswaRegister" as const,
  },
  {
    id: "in-house" as const,
    label: "In-House",
    listKey: "inHouseCertificates" as const,
  },
];

export const PortalShell: React.FC<PortalShellProps> = (props) => {
  const { mode, spHttpClient, webUrl, userEmail } = props;
  const client = React.useMemo(
    () => ({ spHttpClient, webUrl }),
    [spHttpClient, webUrl]
  );

  const [permission, setPermission] = React.useState<PermissionProfile | null>(
    null
  );
  const [permError, setPermError] = React.useState<string | null>(null);
  const [loadingPerm, setLoadingPerm] = React.useState(true);

  const [view, setView] = React.useState<AdminViewId | CustomerViewId>(
    "dashboard"
  );
  const [recordTab, setRecordTab] = React.useState<
    "npors" | "eusr" | "streetworks" | "in-house"
  >("npors");
  const [rows, setRows] = React.useState<PortalTableRow[]>([]);
  const [headers, setHeaders] = React.useState<string[]>([]);
  const [counts, setCounts] = React.useState<Record<string, number>>({});
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  const nav = mode === "admin" ? ADMIN_NAV : CUSTOMER_NAV;
  const views: Record<string, ViewSpec> =
    mode === "admin" ? ADMIN_VIEWS : CUSTOMER_VIEWS;

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoadingPerm(true);
      setPermError(null);
      try {
        const profile = await getActivePermissionByEmail(client, userEmail);
        if (!cancelled) {
          setPermission(profile);
        }
      } catch (e) {
        if (!cancelled) {
          setPermError(
            e instanceof Error ? e.message : "Failed to load permissions."
          );
        }
      } finally {
        if (!cancelled) {
          setLoadingPerm(false);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [client, userEmail]);

  React.useEffect(() => {
    if (!permission) {
      return;
    }
    if (mode === "admin" && permission.roleType !== "Admin") {
      return;
    }
    if (mode === "customer" && permission.roleType !== "Customer") {
      return;
    }

    let cancelled = false;
    (async () => {
      setLoading(true);
      setError(null);
      try {
        const companyName =
          mode === "customer" ? permission.companyDisplayName : undefined;

        if (view === "dashboard") {
          const c = await loadDashboardCounts(client, companyName);
          if (!cancelled) {
            setCounts(c);
            setRows([]);
            setHeaders([]);
          }
          return;
        }

        if (view === "training-records") {
          const tab = RECORD_SUBNAV.filter((t) => t.id === recordTab)[0];
          const fields = getSharePointFields(tab.listKey);
          const cols =
            tab.listKey === "nporsRegister"
              ? [
                  fields.candidateName,
                  fields.nporsCategory,
                  fields.expiry,
                  fields.trainingOutcome,
                ]
              : tab.listKey === "eusrRegister"
              ? [
                  fields.candidateName,
                  fields.eusrCategory,
                  fields.expiry,
                  fields.trainingOutcome,
                ]
              : tab.listKey === "nrswaRegister"
              ? [
                  fields.candidateName,
                  fields.course,
                  fields.expiryDate,
                  fields.trainingOutcome,
                ]
              : [
                  fields.candidateName,
                  fields.courseCategory,
                  fields.expiryDate,
                  fields.trainingOutcome,
                ];
          const data = await loadPortalListRows(client, tab.listKey, {
            columns: cols,
            companyName:
              mode === "customer" ? permission.companyDisplayName : undefined,
            companyFieldInternalName:
              mode === "customer" ? fields.companyName : undefined,
            customerVisibleOnly: mode === "customer",
            visibleFieldInternalName:
              mode === "customer" ? fields.customerVisible : undefined,
          });
          if (!cancelled) {
            setHeaders(["Candidate", "Category / course", "Expiry", "Outcome"]);
            setRows(data);
          }
          return;
        }

        const spec = views[view];
        if (!spec || spec.stub || !spec.listKey || !spec.columns || !spec.headers) {
          if (!cancelled) {
            setRows([]);
            setHeaders([]);
          }
          return;
        }

        const data = await loadPortalListRows(client, spec.listKey, {
          columns: spec.columns,
          companyName:
            mode === "customer" && spec.customerScoped
              ? permission.companyDisplayName
              : undefined,
          companyFieldInternalName: spec.companyField,
          customerVisibleOnly: Boolean(
            mode === "customer" && spec.visibleField
          ),
          visibleFieldInternalName: spec.visibleField,
        });
        if (!cancelled) {
          setHeaders(spec.headers);
          setRows(data);
        }
      } catch (e) {
        if (!cancelled) {
          setError(e instanceof Error ? e.message : "Failed to load data.");
          setRows([]);
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [permission, view, recordTab, mode, client, views]);

  if (loadingPerm) {
    return <div className={styles.muted}>Checking permissions…</div>;
  }

  if (permError) {
    return <div className={styles.error}>{permError}</div>;
  }

  if (!permission) {
    return (
      <div className={styles.panel}>
        <h2 className={styles.title}>Access denied</h2>
        <p className={styles.muted}>
          No active Permissions List row for {userEmail}. Add an Active Training
          Manager (Admin) or Supervisor (Customer) entry in SharePoint.
        </p>
      </div>
    );
  }

  if (mode === "admin" && permission.roleType !== "Admin") {
    return (
      <div className={styles.panel}>
        <h2 className={styles.title}>Admin access required</h2>
        <p className={styles.muted}>
          Signed in as {permission.userEmail} with role Customer. Use the
          Customer Portal web part, or grant Training Manager in Permissions
          List.
        </p>
      </div>
    );
  }

  if (mode === "customer" && permission.roleType !== "Customer") {
    return (
      <div className={styles.panel}>
        <h2 className={styles.title}>Customer access required</h2>
        <p className={styles.muted}>
          Signed in as {permission.userEmail} with role Admin. Use the Admin
          Portal web part for training managers.
        </p>
      </div>
    );
  }

  const spec = views[view];

  return (
    <div className={styles.shell}>
      <aside className={styles.sidebar}>
        <p className={styles.brand}>PAVE</p>
        <p className={styles.tagline}>Paving the way in industry</p>
        {mode === "customer" && permission.companyDisplayName ? (
          <span className={styles.chip}>{permission.companyDisplayName}</span>
        ) : (
          <span className={styles.chip}>
            {mode === "admin" ? "Admin operations" : "Customer portal"}
          </span>
        )}
        {nav.map((item) => (
          <button
            key={item.id}
            type="button"
            className={`${styles.navBtn} ${
              view === item.id ? styles.navBtnActive : ""
            }`}
            onClick={() => setView(item.id)}
          >
            {item.label}
          </button>
        ))}
      </aside>
      <main className={styles.main}>
        <h1 className={styles.title}>{spec.title}</h1>
        <p className={styles.subtitle}>{spec.subtitle}</p>

        {view === "dashboard" && (
          <div className={styles.stats}>
            {Object.keys(counts).map((key) => (
              <div key={key} className={styles.stat}>
                <p className={styles.statValue}>{counts[key]}</p>
                <p className={styles.statLabel}>{key}</p>
              </div>
            ))}
          </div>
        )}

        {view === "training-records" && (
          <div className={styles.toolbar}>
            {RECORD_SUBNAV.map((tab) => (
              <button
                key={tab.id}
                type="button"
                className={
                  recordTab === tab.id ? styles.toolbarBtnActive : undefined
                }
                onClick={() => setRecordTab(tab.id)}
              >
                {tab.label}
              </button>
            ))}
          </div>
        )}

        {spec.stub ? (
          <div className={styles.panel}>
            <p className={styles.muted}>{spec.stub}</p>
          </div>
        ) : view !== "dashboard" ? (
          <div className={styles.panel}>
            {loading && <p className={styles.muted}>Loading…</p>}
            {error && <p className={styles.error}>{error}</p>}
            {!loading && !error && (
              <table className={styles.table}>
                <thead>
                  <tr>
                    {headers.map((h) => (
                      <th key={h}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={headers.length || 1}
                        className={styles.muted}
                      >
                        No rows found.
                      </td>
                    </tr>
                  ) : (
                    rows.map((row) => (
                      <tr key={row.id}>
                        {row.cells.map((c, i) => (
                          <td key={`${row.id}-${i}`}>{c}</td>
                        ))}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            )}
          </div>
        ) : null}
      </main>
    </div>
  );
};

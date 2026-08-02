import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";

import { getSharePointFields } from "../schema/sharepointSchema";
import {
  documentRowsToPortalTable,
  filterPortalRowsByAccess,
  loadCustomerDocuments,
} from "../services/customerAccessService";
import {
  getActivePermissionByEmail,
  siteAdminPermissionProfile,
} from "../services/permissionService";
import {
  getAdminSchemaColumns,
  loadCustomerDashboardData,
  loadDashboardCounts,
  loadPortalListRows,
  type CustomerDashboardData,
  type PortalTableRow,
} from "../services/portalDataService";
import { normalizeSharePointUserEmail } from "../services/sharePointListService";
import type { PermissionProfile } from "../types/models";
import { AdminBulkUpload } from "./AdminBulkUpload";
import { AdminDataTable } from "./AdminDataTable";
import { AdminHubDashboard } from "./AdminHubDashboard";
import { CustomerPortalView } from "./CustomerPortalView";
import {
  ADMIN_NAV,
  CUSTOMER_NAV,
  type AdminViewId,
  type CustomerViewId,
} from "./nav";
import styles from "./portal.module.scss";
import type { SharePointListKey } from "../schema/sharepointSchema";

const paveLogo: string = require("../assets/pave-logo.png");

export type PortalMode = "admin" | "customer";

export interface PortalShellProps {
  mode: PortalMode;
  spHttpClient: SPHttpClient;
  webUrl: string;
  userEmail: string;
  /** SharePoint site collection admin — can open Admin even without Training Manager row. */
  isSiteAdmin?: boolean;
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
    subtitle: "Training Matrix Update — wide expiry grid (same list as Next.js admin)",
    listKey: "trainingMatrix",
    headers: [
      "Name",
      "DOB",
      "CSCS Expiry",
      "SSSTS Expiry",
      "SMSTS Expiry",
      "NRSWA Expiry",
      "EUSR Expiry",
    ],
    columns: [
      "Title",
      "DOB",
      "CSCSExpiry",
      "SSSTSExpiry",
      "SMSTSExpiry",
      "NRSWAExpiry",
      "EUSRExpiry",
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
  "bulk-upload": {
    title: "Bulk upload",
    subtitle: "Import CSV into SharePoint lists (same types as Next.js admin)",
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
    subtitle: "Your company's matrix rows (Training Matrix Update)",
    listKey: "trainingMatrix",
    headers: ["Name", "CSCS Expiry", "EUSR Expiry", "NRSWA Expiry"],
    columns: ["Title", "CSCSExpiry", "EUSRExpiry", "NRSWAExpiry"],
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
    subtitle: "Certificates, card scans, NVQs, brochures and shared files",
    listKey: "customerDocuments",
    headers: [
      "Document Name",
      "Document Type",
      "Candidate Name",
      "Modified Date",
      "View",
      "Download",
    ],
    columns: [
      "FileLeafRef",
      "DocumentType",
      "Candidate",
      "Modified",
      "__view",
      "__download",
    ],
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
  const { mode, spHttpClient, webUrl, userEmail, isSiteAdmin } = props;
  const client = React.useMemo(
    () => ({ spHttpClient, webUrl }),
    [spHttpClient, webUrl]
  );
  const resolvedEmail = normalizeSharePointUserEmail(userEmail);

  const [permission, setPermission] = React.useState<PermissionProfile | null>(
    null
  );
  const [permError, setPermError] = React.useState<string | null>(null);
  const [loadingPerm, setLoadingPerm] = React.useState(true);

  const [view, setView] = React.useState<AdminViewId | CustomerViewId>(
    mode === "customer" ? "training-matrix" : "dashboard"
  );
  const [recordTab, setRecordTab] = React.useState<
    "npors" | "eusr" | "streetworks" | "in-house"
  >("npors");
  const [rows, setRows] = React.useState<PortalTableRow[]>([]);
  const [headers, setHeaders] = React.useState<string[]>([]);
  const [columns, setColumns] = React.useState<string[]>([]);
  const [activeListKey, setActiveListKey] =
    React.useState<SharePointListKey | null>(null);
  const [counts, setCounts] = React.useState<Record<string, number>>({});
  const [customerDash, setCustomerDash] =
    React.useState<CustomerDashboardData | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [refreshTick, setRefreshTick] = React.useState(0);
  const [mobileNavOpen, setMobileNavOpen] = React.useState(false);

  const nav = mode === "admin" ? ADMIN_NAV : CUSTOMER_NAV;
  const views: Record<string, ViewSpec> =
    mode === "admin" ? ADMIN_VIEWS : CUSTOMER_VIEWS;

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoadingPerm(true);
      setPermError(null);
      try {
        let profile = await getActivePermissionByEmail(
          client,
          resolvedEmail || userEmail
        );

        // Site collection admins can always use the Admin web part.
        if (
          mode === "admin" &&
          isSiteAdmin &&
          (!profile || !profile.canAccessAdmin)
        ) {
          profile = siteAdminPermissionProfile(resolvedEmail || userEmail);
        }

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
  }, [client, userEmail, resolvedEmail, mode, isSiteAdmin]);

  React.useEffect(() => {
    if (!permission) {
      return;
    }
    if (mode === "admin" && !permission.canAccessAdmin) {
      return;
    }
    if (mode === "customer" && !permission.canAccessCustomer) {
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
          if (mode === "customer" && permission.companyDisplayName) {
            const dash = await loadCustomerDashboardData(
              client,
              permission.companyDisplayName,
              permission.companyId
            );
            if (!cancelled) {
              const scopedMatrix = filterPortalRowsByAccess(
                dash.matrixRows,
                permission,
                { candidateNameField: "Title" }
              );
              setCustomerDash({
                ...dash,
                matrixRows: scopedMatrix,
                counts: {
                  ...dash.counts,
                  workforce: scopedMatrix.length,
                  matrix: scopedMatrix.length,
                },
              });
              setCounts({
                ...dash.counts,
                workforce: scopedMatrix.length,
                matrix: scopedMatrix.length,
              });
              setRows(scopedMatrix);
              setHeaders(["Name", "Status", "CSCS Expiry"]);
              setColumns(["Title", "Status", "CSCSExpiry"]);
              setActiveListKey(null);
            }
          } else {
            const c = await loadDashboardCounts(client, companyName);
            if (!cancelled) {
              setCustomerDash(null);
              setCounts(c);
              setRows([]);
              setHeaders([]);
              setColumns([]);
              setActiveListKey(null);
            }
          }
          return;
        }

        if (view === "training-records") {
          const tab = RECORD_SUBNAV.filter((t) => t.id === recordTab)[0];
          let cols: string[];
          let hdrs: string[];
          let companyField: string = "CompanyName";
          let visibleField: string = "CustomerVisible";

          if (mode === "admin") {
            const schema = getAdminSchemaColumns(tab.listKey);
            cols = schema.columns;
            hdrs = schema.headers;
            const fields = getSharePointFields(tab.listKey);
            companyField = fields.companyName;
            visibleField = fields.customerVisible;
          } else if (tab.listKey === "nporsRegister") {
            const fields = getSharePointFields("nporsRegister");
            cols = [
              fields.candidateName,
              fields.nporsCategory,
              fields.expiry,
              fields.trainingOutcome,
            ];
            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
            companyField = fields.companyName;
            visibleField = fields.customerVisible;
          } else if (tab.listKey === "eusrRegister") {
            const fields = getSharePointFields("eusrRegister");
            cols = [
              fields.candidateName,
              fields.eusrCategory,
              fields.expiry,
              fields.trainingOutcome,
            ];
            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
            companyField = fields.companyName;
            visibleField = fields.customerVisible;
          } else if (tab.listKey === "nrswaRegister") {
            const fields = getSharePointFields("nrswaRegister");
            cols = [
              fields.candidateName,
              fields.course,
              fields.expiryDate,
              fields.trainingOutcome,
            ];
            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
            companyField = fields.companyName;
            visibleField = fields.customerVisible;
          } else {
            const fields = getSharePointFields("inHouseCertificates");
            cols = [
              fields.candidateName,
              fields.courseCategory,
              fields.expiryDate,
              fields.trainingOutcome,
            ];
            hdrs = ["Candidate", "Category / course", "Expiry", "Outcome"];
            companyField = fields.companyName;
            visibleField = fields.customerVisible;
          }

          const data = await loadPortalListRows(client, tab.listKey, {
            columns: cols,
            companyName:
              mode === "customer" ? permission.companyDisplayName : undefined,
            companyFieldInternalName:
              mode === "customer" ? companyField : undefined,
            customerVisibleOnly: mode === "customer",
            visibleFieldInternalName:
              mode === "customer" ? visibleField : undefined,
            loadAll: mode === "admin",
          });
          const scoped =
            mode === "customer"
              ? filterPortalRowsByAccess(data, permission)
              : data;
          if (!cancelled) {
            setHeaders(hdrs);
            setColumns(cols);
            setActiveListKey(tab.listKey);
            setRows(scoped);
          }
          return;
        }

        const spec = views[view];
        if (!spec || spec.stub || !spec.listKey) {
          if (!cancelled) {
            setRows([]);
            setHeaders([]);
            setColumns([]);
            setActiveListKey(null);
          }
          return;
        }

        // Customer documents — resolve candidate names + view/download URLs.
        if (mode === "customer" && spec.listKey === "customerDocuments") {
          const docsRows = await loadCustomerDocuments(client, permission);
          if (!cancelled) {
            setHeaders(spec.headers || []);
            setColumns(spec.columns || []);
            setActiveListKey("customerDocuments");
            setRows(documentRowsToPortalTable(docsRows));
          }
          return;
        }

        let cols = spec.columns;
        let headerLabels = spec.headers;
        // Admin tables: use full schema except Training Matrix Update (too wide —
        // keep curated display columns; edit form still gets full schema fields).
        if (mode === "admin" && spec.listKey && spec.listKey !== "trainingMatrix") {
          const schema = getAdminSchemaColumns(spec.listKey);
          cols = schema.columns;
          headerLabels = schema.headers;
        }

        if (!cols || !headerLabels) {
          if (!cancelled) {
            setRows([]);
            setHeaders([]);
            setColumns([]);
            setActiveListKey(null);
          }
          return;
        }

        // Training Matrix Update has no company column — scope via Workforce names.
        if (
          mode === "customer" &&
          spec.listKey === "trainingMatrix" &&
          permission.companyDisplayName
        ) {
          const workforce = getSharePointFields("workforce");
          const wf = await loadPortalListRows(client, "workforce", {
            columns: [workforce.candidateName],
            companyName: permission.companyDisplayName,
            companyFieldInternalName: workforce.companyName,
            loadAll: true,
          });
          const nameSet: { [name: string]: boolean } = {};
          for (let i = 0; i < wf.length; i++) {
            const n = (wf[i].cells[0] || "").trim().toLowerCase();
            if (n) nameSet[n] = true;
          }
          const allMatrix = await loadPortalListRows(client, "trainingMatrix", {
            columns: cols,
            loadAll: true,
          });
          const scopedMatrix = allMatrix.filter((row) => {
            const title = (row.cells[0] || "").trim().toLowerCase();
            return Boolean(title && nameSet[title]);
          });
          const accessFiltered = filterPortalRowsByAccess(scopedMatrix, permission, {
            candidateNameField: "Title",
          });
          if (!cancelled) {
            setHeaders(headerLabels);
            setColumns(cols);
            setActiveListKey("trainingMatrix");
            setRows(accessFiltered);
          }
          return;
        }

        const useLookupId =
          mode === "customer" &&
          (spec.listKey === "events" || spec.listKey === "customerDocuments") &&
          permission.companyId &&
          permission.companyId !== "0";

        const data = await loadPortalListRows(client, spec.listKey, {
          columns: cols,
          companyName:
            mode === "customer" && spec.customerScoped && !useLookupId
              ? permission.companyDisplayName
              : undefined,
          companyFieldInternalName:
            mode === "customer" && !useLookupId ? spec.companyField : undefined,
          companyId: useLookupId ? permission.companyId : undefined,
          companyIdFieldInternalName: useLookupId
            ? spec.listKey === "events"
              ? "EventCompanyId"
              : "CompanyId"
            : undefined,
          customerVisibleOnly: Boolean(
            mode === "customer" && spec.visibleField
          ),
          visibleFieldInternalName: spec.visibleField,
          loadAll: mode === "admin",
        });

        const scoped =
          mode === "customer" &&
          (spec.listKey === "workforce" ||
            spec.listKey === "trainingMatrix" ||
            spec.listKey === "nvqRegister")
            ? filterPortalRowsByAccess(data, permission, {
                candidateNameField:
                  spec.listKey === "trainingMatrix" ? "Title" : undefined,
              })
            : data;

        if (!cancelled) {
          setHeaders(headerLabels);
          setColumns(cols);
          setActiveListKey(spec.listKey);
          setRows(scoped);
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
  }, [permission, view, recordTab, mode, client, views, refreshTick]);

  React.useEffect(() => {
    setMobileNavOpen(false);
  }, [view]);

  React.useEffect(() => {
    if (!mobileNavOpen) {
      return;
    }
    const previous = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = previous;
    };
  }, [mobileNavOpen]);

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
          No active Permissions List row for {resolvedEmail || userEmail}. Add
          an Active Admin / Training Manager / Supervisor / Candidate entry in
          SharePoint Permissions List with matching User Email.
        </p>
      </div>
    );
  }

  if (mode === "admin" && !permission.canAccessAdmin) {
    return (
      <div className={styles.panel}>
        <h2 className={styles.title}>Admin access required</h2>
        <p className={styles.muted}>
          Signed in as {permission.userEmail} ({permission.roleLabel}). This
          account can use the Customer Portal web part, but not Admin.
        </p>
        <p className={styles.muted}>
          PAVE internal Admin needs Role Type = Admin (or Training Manager for
          legacy access). Site collection admins can also open Admin after
          refresh.
        </p>
      </div>
    );
  }

  if (mode === "customer" && !permission.canAccessCustomer) {
    return (
      <div className={styles.panel}>
        <h2 className={styles.title}>Customer access required</h2>
        <p className={styles.muted}>
          Signed in as {permission.userEmail} with role {permission.roleLabel}.
          Use the Admin Portal web part for PAVE internal admin accounts.
        </p>
      </div>
    );
  }

  const spec = views[view];

  // Customer portal — design presentation layer
  if (mode === "customer") {
    if (!permission.companyDisplayName) {
      return (
        <div className={styles.panel}>
          <h2 className={styles.title}>Company not resolved</h2>
          <p className={styles.muted}>
            Your Permissions List row is Active, but the Company lookup could
            not be resolved to a company name (CompanyId{" "}
            {permission.companyId}). Check the Company field on your Permissions
            List item.
          </p>
        </div>
      );
    }
    return (
      <CustomerPortalView
        permission={permission}
        view={view as CustomerViewId}
        onNavigate={(id) => setView(id)}
        counts={counts}
        headers={headers}
        rows={rows}
        loading={loading}
        error={error}
        stub={spec.stub}
        pageTitle={spec.title}
        pageSubtitle={spec.subtitle}
        dashboard={customerDash}
      />
    );
  }

  return (
    <div className={`${styles.shell} ${styles.adminShell}`}>
      <header className={styles.adminTopNav}>
        <div className={styles.adminTopNavBar}>
          <div className={styles.adminBrandBlock}>
            <img
              className={styles.adminBrandLogo}
              src={paveLogo}
              alt="PAVE Training"
            />
            <p className={styles.tagline}>Admin operations</p>
          </div>

          <nav className={styles.adminNavDesktop} aria-label="Admin">
            {nav.map((item) => (
              <button
                key={item.id}
                type="button"
                className={`${styles.adminNavLink} ${
                  view === item.id ? styles.adminNavLinkActive : ""
                }`}
                onClick={() => setView(item.id)}
              >
                {item.label}
              </button>
            ))}
          </nav>

          <div className={styles.adminTopNavTrailing}>
            <span className={styles.adminUserChip} title={permission.userEmail}>
              {permission.userEmail}
            </span>
            <button
              type="button"
              className={styles.adminMenuToggle}
              aria-expanded={mobileNavOpen}
              aria-controls="pave-admin-mobile-nav"
              onClick={() => setMobileNavOpen((open) => !open)}
            >
              {mobileNavOpen ? "Close" : "Menu"}
            </button>
          </div>
        </div>

        {mobileNavOpen ? (
          <div
            className={styles.adminMobileScrim}
            onClick={() => setMobileNavOpen(false)}
            aria-hidden
          />
        ) : null}

        <div
          id="pave-admin-mobile-nav"
          className={`${styles.adminMobileDrawer} ${
            mobileNavOpen ? styles.adminMobileDrawerOpen : ""
          }`}
          hidden={!mobileNavOpen}
        >
          <p className={styles.adminMobileTitle}>All options</p>
          <nav className={styles.adminMobileNav} aria-label="Admin mobile">
            {nav.map((item) => (
              <button
                key={item.id}
                type="button"
                className={`${styles.adminMobileLink} ${
                  view === item.id ? styles.adminMobileLinkActive : ""
                }`}
                onClick={() => {
                  setView(item.id);
                  setMobileNavOpen(false);
                }}
              >
                {item.label}
              </button>
            ))}
          </nav>
        </div>
      </header>

      <main className={styles.main}>
        {view === "dashboard" ? (
          <AdminHubDashboard
            userEmail={permission.userEmail}
            counts={counts}
            onNavigate={(id) => setView(id)}
          />
        ) : (
          <>
            <header className={styles.pageHeader}>
              <div>
                <p className={styles.eyebrow}>Admin</p>
                <h1 className={styles.title}>{spec.title}</h1>
                <p className={styles.subtitle}>
                  {spec.subtitle}
                  {!spec.stub && !loading ? " · " + rows.length + " rows" : ""}
                </p>
              </div>
            </header>

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
            ) : view === "bulk-upload" ? (
              <AdminBulkUpload client={client} />
            ) : activeListKey ? (
              <AdminDataTable
                client={client}
                listKey={activeListKey}
                title={spec.title}
                headers={headers}
                columns={columns}
                rows={rows}
                loading={loading}
                error={error}
                onRefresh={() => setRefreshTick((n) => n + 1)}
                formColumns={
                  activeListKey === "trainingMatrix"
                    ? getAdminSchemaColumns("trainingMatrix").columns
                    : undefined
                }
              />
            ) : (
              <div className={styles.panel}>
                {loading && <p className={styles.muted}>Loading…</p>}
                {error && <p className={styles.error}>{error}</p>}
              </div>
            )}
          </>
        )}
      </main>
    </div>
  );
};

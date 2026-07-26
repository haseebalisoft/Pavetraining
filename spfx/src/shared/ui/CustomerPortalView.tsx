import * as React from "react";

import type {
  CustomerDashboardData,
  PortalTableRow,
} from "../services/portalDataService";
import { accessScopeBadgeLabel } from "../services/permissionService";
import type { PermissionProfile } from "../types/models";
import type { CustomerViewId } from "./nav";
import styles from "./customerPortal.module.scss";

// eslint-disable-next-line @typescript-eslint/no-var-requires
const paveLogo: string = require("../assets/pave-logo.png");

export interface CustomerPortalViewProps {
  permission: PermissionProfile;
  view: CustomerViewId;
  onNavigate: (view: CustomerViewId) => void;
  counts: Record<string, number>;
  headers: string[];
  rows: PortalTableRow[];
  loading: boolean;
  error: string | null;
  stub?: string;
  pageTitle: string;
  pageSubtitle: string;
  /** Real SharePoint dashboard bundle (matrix / docs / nvq / events / offers). */
  dashboard?: CustomerDashboardData | null;
}

const DESKTOP_NAV: ReadonlyArray<{ id: CustomerViewId; label: string }> = [
  { id: "training-matrix", label: "Training Matrix" },
  { id: "dashboard", label: "Dashboard" },
  { id: "candidates", label: "Candidates" },
  { id: "documents", label: "Documents" },
  { id: "nvq-progress", label: "NVQ Progress" },
  { id: "events", label: "Events" },
  { id: "offers", label: "Offers" },
  { id: "support", label: "Support" },
];

const MORE_NAV: ReadonlyArray<{ id: CustomerViewId; label: string }> = [
  { id: "training-records", label: "Training Records" },
  { id: "candidates", label: "Candidates" },
  { id: "offers", label: "Offers" },
  { id: "support", label: "Support" },
];

function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
  return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
}

function statusKind(
  raw: string
): "ok" | "warn" | "bad" {
  const s = (raw || "").toLowerCase();
  if (
    s.indexOf("missing") >= 0 ||
    s.indexOf("expired") >= 0 ||
    s.indexOf("review") >= 0 ||
    s.indexOf("fail") >= 0
  ) {
    return "bad";
  }
  if (
    s.indexOf("expir") >= 0 ||
    s.indexOf("soon") >= 0 ||
    s.indexOf("due") >= 0 ||
    s.indexOf("attention") >= 0
  ) {
    return "warn";
  }
  return "ok";
}

function statusLabel(raw: string): string {
  const kind = statusKind(raw);
  if (kind === "bad") return raw.trim() || "Missing Data";
  if (kind === "warn") return raw.trim() || "Expiring Soon";
  if (!raw.trim() || raw.toLowerCase() === "compliant") return "Compliant";
  return raw.trim();
}

function IconPeople(): React.ReactElement {
  return (
    <svg className={styles.navIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M16 11a4 4 0 1 0-4-4 4 4 0 0 0 4 4ZM8 13a4 4 0 1 0-4-4 4 4 0 0 0 4 4ZM16 13c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4ZM8 13c-.29 0-.62.02-.97.05C4.84 13.56 2 14.94 2 17v2h6"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function IconMatrix(): React.ReactElement {
  return (
    <svg className={styles.navIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M4 4h7v7H4V4Zm9 0h7v7h-7V4ZM4 13h7v7H4v-7Zm9 0h7v7h-7v-7Z"
        stroke="currentColor"
        strokeWidth="1.7"
      />
    </svg>
  );
}

function IconDocs(): React.ReactElement {
  return (
    <svg className={styles.navIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M7 3h7l5 5v13a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"
        stroke="currentColor"
        strokeWidth="1.7"
      />
      <path d="M14 3v5h5" stroke="currentColor" strokeWidth="1.7" />
    </svg>
  );
}

function IconNvq(): React.ReactElement {
  return (
    <svg className={styles.navIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M12 3 4 7v5c0 5 3.4 8.4 8 9 4.6-.6 8-4 8-9V7l-8-4Z"
        stroke="currentColor"
        strokeWidth="1.7"
      />
      <path d="m9 12 2 2 4-4" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconEvents(): React.ReactElement {
  return (
    <svg className={styles.navIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <rect x="3" y="5" width="18" height="16" rx="2" stroke="currentColor" strokeWidth="1.7" />
      <path d="M3 10h18M8 3v4M16 3v4" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconOffers(): React.ReactElement {
  return (
    <svg className={styles.navIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M12 3 4 7v4c0 4.5 3 8.2 8 10 5-1.8 8-5.5 8-10V7l-8-4Z"
        stroke="currentColor"
        strokeWidth="1.7"
      />
      <path d="M9 12h6" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconSupport(): React.ReactElement {
  return (
    <svg className={styles.navIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.7" />
      <path
        d="M9.5 9.5a2.5 2.5 0 1 1 3.6 2.2c-.7.4-1.1.9-1.1 1.8M12 17h.01"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinecap="round"
      />
    </svg>
  );
}

function IconClock(): React.ReactElement {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="8" stroke="currentColor" strokeWidth="1.7" />
      <path d="M12 8v5l3 2" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconWarn(): React.ReactElement {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M12 4 3 19h18L12 4Z"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinejoin="round"
      />
      <path d="M12 10v4M12 16.5h.01" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function IconFolder(): React.ReactElement {
  return (
    <svg className={styles.docIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7Z"
        fill="#F5C542"
        stroke="#D4A017"
        strokeWidth="1"
      />
    </svg>
  );
}

function IconPdf(): React.ReactElement {
  return (
    <svg className={styles.docIcon} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M7 3h7l5 5v13H7V3Z" fill="#FEE2E2" stroke="#DC2626" strokeWidth="1.4" />
      <path d="M14 3v5h5" stroke="#DC2626" strokeWidth="1.4" />
      <text x="8" y="17" fontSize="6" fontWeight="700" fill="#DC2626">
        PDF
      </text>
    </svg>
  );
}

function IconPin(): React.ReactElement {
  return (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M12 21s7-5.2 7-11a7 7 0 1 0-14 0c0 5.8 7 11 7 11Z"
        stroke="currentColor"
        strokeWidth="1.7"
      />
      <circle cx="12" cy="10" r="2.2" stroke="currentColor" strokeWidth="1.7" />
    </svg>
  );
}

function navIcon(id: CustomerViewId): React.ReactNode {
  switch (id) {
    case "dashboard":
      return <IconPeople />;
    case "training-matrix":
    case "candidates":
      return <IconMatrix />;
    case "documents":
      return <IconDocs />;
    case "nvq-progress":
      return <IconNvq />;
    case "events":
      return <IconEvents />;
    case "offers":
      return <IconOffers />;
    case "support":
      return <IconSupport />;
    default:
      return <IconMatrix />;
  }
}

function parseEventDate(raw: string): { month: string; day: string } {
  const d = new Date(raw);
  if (!isNaN(d.getTime())) {
    return {
      month: d.toLocaleString("en-GB", { month: "short" }).toUpperCase(),
      day: String(d.getDate()),
    };
  }
  return { month: "—", day: "—" };
}

/**
 * Presentational Customer Portal shell — matches PAVE design mock.
 * Receives already-loaded portal state; does not fetch.
 */
export const CustomerPortalView: React.FC<CustomerPortalViewProps> = (props) => {
  const {
    permission,
    view,
    onNavigate,
    counts,
    headers,
    rows,
    loading,
    error,
    stub,
    pageTitle,
    pageSubtitle,
    dashboard,
  } = props;

  const [menuOpen, setMenuOpen] = React.useState(false);
  const [matrixQuery, setMatrixQuery] = React.useState("");

  const companyName =
    permission.companyDisplayName || permission.userEmail || "Customer";
  const avatar = initials(companyName);
  const accessLabel = accessScopeBadgeLabel(permission);
  const downloadLabel = permission.canDownload
    ? "Downloads: Enabled"
    : "Downloads: Disabled";

  const totalCandidates =
    counts.workforce != null
      ? counts.workforce
      : dashboard && dashboard.counts.workforce != null
        ? dashboard.counts.workforce
        : 0;
  const upcomingEvents =
    counts.events != null
      ? counts.events
      : dashboard && dashboard.counts.events != null
        ? dashboard.counts.events
        : 0;
  const expiringSoon =
    counts.expiringSoon != null
      ? counts.expiringSoon
      : dashboard && dashboard.counts.expiringSoon != null
        ? dashboard.counts.expiringSoon
        : 0;
  const missingData =
    counts.missingData != null
      ? counts.missingData
      : dashboard && dashboard.counts.missingData != null
        ? dashboard.counts.missingData
        : 0;

  const sourceMatrix =
    dashboard && dashboard.matrixRows.length > 0
      ? dashboard.matrixRows
      : rows;

  const matrixPreview = React.useMemo(() => {
    const q = matrixQuery.trim().toLowerCase();
    let list = sourceMatrix;
    if (q) {
      list = list.filter(
        (r) => r.cells.join(" ").toLowerCase().indexOf(q) >= 0
      );
    }
    return list.slice(0, 5);
  }, [sourceMatrix, matrixQuery]);

  const closeMenu = (): void => setMenuOpen(false);

  const go = (id: CustomerViewId): void => {
    onNavigate(id);
    closeMenu();
  };

  return (
    <div className={styles.root}>
      <header className={styles.header}>
        <div className={styles.headerLeft}>
          <img
            className={styles.logo}
            src={paveLogo}
            alt="PAVE Training"
          />
          <div className={styles.headerCopy}>
            <h1 className={styles.portalTitle}>PAVE Training Customer Portal</h1>
            <p className={styles.welcome}>
              Welcome back,{" "}
              <span className={styles.welcomeName}>{permission.userEmail}</span>
            </p>
            <p className={styles.tagline}>
              Company: {companyName}
            </p>
            <div className={styles.accessBadges} aria-label="Access summary">
              <span className={styles.badge}>Role: {permission.roleLabel}</span>
              <span className={styles.badge}>Access: {accessLabel}</span>
              <span
                className={
                  permission.canDownload
                    ? styles.badgeOk
                    : styles.badgeMuted
                }
              >
                {downloadLabel}
              </span>
            </div>
          </div>
          <button
            type="button"
            className={styles.menuToggle}
            aria-label="Open menu"
            onClick={() => setMenuOpen(!menuOpen)}
          >
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
              <path
                d="M4 7h16M4 12h16M4 17h16"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>
        <button type="button" className={styles.userMenu} onClick={() => undefined}>
          <span className={styles.userName}>{permission.userEmail}</span>
          <span className={styles.chevron} aria-hidden="true">
            ▾
          </span>
          <span className={styles.avatar}>{avatar}</span>
        </button>
      </header>

      {menuOpen && (
        <div className={styles.mobileNavDrawer}>
          {DESKTOP_NAV.concat(MORE_NAV.filter((m) => m.id === "candidates" || m.id === "training-records")).map(
            (item) => (
              <button
                key={"m-" + item.id}
                type="button"
                className={`${styles.navItem} ${
                  view === item.id ? styles.navItemActive : ""
                }`}
                onClick={() => go(item.id)}
              >
                {navIcon(item.id)}
                {item.label}
              </button>
            )
          )}
        </div>
      )}

      <nav className={styles.navBar} aria-label="Customer portal">
        {DESKTOP_NAV.map((item) => (
          <button
            key={item.id}
            type="button"
            className={`${styles.navItem} ${
              view === item.id ? styles.navItemActive : ""
            }`}
            onClick={() => go(item.id)}
          >
            {navIcon(item.id)}
            {item.label}
          </button>
        ))}
      </nav>

      <div className={styles.body}>
        <p className={styles.mobileWelcome}>
          Welcome back,{" "}
          <span className={styles.welcomeName}>{companyName}</span>
        </p>

        {view === "dashboard" ? (
          <DashboardBody
            totalCandidates={totalCandidates}
            expiringSoon={expiringSoon}
            missingData={missingData}
            upcomingEvents={upcomingEvents}
            matrixRows={matrixPreview}
            matrixTotal={sourceMatrix.length || totalCandidates}
            matrixQuery={matrixQuery}
            onMatrixQuery={setMatrixQuery}
            onNavigate={go}
            loading={loading}
            error={error}
            documentTiles={
              dashboard && dashboard.documentTiles.length > 0
                ? dashboard.documentTiles
                : []
            }
            nvqRows={
              dashboard && dashboard.nvqRows.length > 0
                ? dashboard.nvqRows
                : []
            }
            eventRows={
              dashboard && dashboard.eventRows.length > 0
                ? dashboard.eventRows
                : []
            }
            offerCards={
              dashboard && dashboard.offerCards.length > 0
                ? dashboard.offerCards
                : []
            }
          />
        ) : (
          <ListBody
            title={pageTitle}
            subtitle={pageSubtitle}
            stub={stub}
            headers={headers}
            rows={rows}
            loading={loading}
            error={error}
            view={view}
            onNavigate={go}
          />
        )}
      </div>

      <nav className={styles.bottomNav} aria-label="Mobile navigation">
        <button
          type="button"
          className={`${styles.tabItem} ${
            view === "dashboard" ? styles.tabItemActive : ""
          }`}
          onClick={() => go("dashboard")}
        >
          {navIcon("dashboard")}
          Dashboard
        </button>
        <button
          type="button"
          className={`${styles.tabItem} ${
            view === "training-matrix" ? styles.tabItemActive : ""
          }`}
          onClick={() => go("training-matrix")}
        >
          {navIcon("training-matrix")}
          Matrix
        </button>
        <button
          type="button"
          className={`${styles.tabItem} ${
            view === "documents" ? styles.tabItemActive : ""
          }`}
          onClick={() => go("documents")}
        >
          {navIcon("documents")}
          Documents
        </button>
        <button
          type="button"
          className={`${styles.tabItem} ${
            view === "events" ? styles.tabItemActive : ""
          }`}
          onClick={() => go("events")}
        >
          {navIcon("events")}
          Events
        </button>
        <button
          type="button"
          className={`${styles.tabItem} ${menuOpen ? styles.tabItemActive : ""}`}
          onClick={() => setMenuOpen(!menuOpen)}
        >
          <svg className={styles.navIcon} viewBox="0 0 24 24" fill="currentColor">
            <circle cx="5" cy="12" r="2" />
            <circle cx="12" cy="12" r="2" />
            <circle cx="19" cy="12" r="2" />
          </svg>
          More
        </button>
      </nav>
    </div>
  );
};

/* Placeholder samples removed — dashboard uses SharePoint data via `dashboard` prop. */

interface DashboardBodyProps {
  totalCandidates: number;
  expiringSoon: number;
  missingData: number;
  upcomingEvents: number;
  matrixRows: PortalTableRow[];
  matrixTotal: number;
  matrixQuery: string;
  onMatrixQuery: (v: string) => void;
  onNavigate: (v: CustomerViewId) => void;
  loading: boolean;
  error: string | null;
  documentTiles: Array<{
    id: string;
    label: string;
    meta: string;
    kind: "folder" | "pdf";
  }>;
  nvqRows: Array<{ id: string; name: string; course: string; pct: number }>;
  eventRows: Array<{
    id: string;
    title: string;
    when: string;
    where: string;
    dateRaw: string;
  }>;
  offerCards: Array<{
    id: string;
    badge: string;
    title: string;
    code: string;
  }>;
}

const DashboardBody: React.FC<DashboardBodyProps> = (p) => {
  return (
    <div>
      <div className={styles.stats}>
        <div className={styles.statCard}>
          <div>
            <p className={styles.statLabel}>Total Candidates</p>
            <p className={styles.statValue}>{p.totalCandidates}</p>
            <p className={`${styles.statHint} ${styles.statHintGood}`}>
              Company workforce
            </p>
          </div>
          <span className={`${styles.statIcon} ${styles.statIconGreen}`}>
            <IconPeople />
          </span>
        </div>
        <div className={styles.statCard}>
          <div>
            <p className={styles.statLabel}>Expiring Soon</p>
            <p className={styles.statValue}>{p.expiringSoon}</p>
            <p className={`${styles.statHint} ${styles.statHintWarn}`}>
              Within 60 days
            </p>
          </div>
          <span className={`${styles.statIcon} ${styles.statIconOrange}`}>
            <IconClock />
          </span>
        </div>
        <div className={styles.statCard}>
          <div>
            <p className={styles.statLabel}>Missing Data</p>
            <p className={styles.statValue}>{p.missingData}</p>
            <p className={`${styles.statHint} ${styles.statHintWarn}`}>
              Requires attention
            </p>
          </div>
          <span className={`${styles.statIcon} ${styles.statIconRed}`}>
            <IconWarn />
          </span>
        </div>
        <div className={styles.statCard}>
          <div>
            <p className={styles.statLabel}>Upcoming Events</p>
            <p className={styles.statValue}>{p.upcomingEvents}</p>
            <p className={`${styles.statHint} ${styles.statHintGood}`}>
              Next 30 days
            </p>
          </div>
          <span className={`${styles.statIcon} ${styles.statIconGreen}`}>
            <IconEvents />
          </span>
        </div>
      </div>

      <div className={styles.dashMain}>
        <section className={styles.card}>
          <div className={styles.cardHeader}>
            <h2 className={styles.cardTitle}>Training Matrix</h2>
            <div className={styles.cardActions}>
              <input
                className={styles.searchInput}
                type="search"
                placeholder="Search candidates..."
                value={p.matrixQuery}
                onChange={(e) => p.onMatrixQuery(e.target.value)}
              />
              <button
                type="button"
                className={styles.ghostBtn}
                onClick={() => undefined}
              >
                Filters
              </button>
              <button
                type="button"
                className={styles.primaryBtn}
                onClick={() => p.onNavigate("training-matrix")}
              >
                View full matrix
              </button>
            </div>
          </div>
          {p.loading && <p className={styles.muted}>Loading…</p>}
          {p.error && <p className={styles.error}>{p.error}</p>}
          {!p.loading && !p.error && (
            <>
              <MatrixTable
                rows={p.matrixRows}
                condensed={false}
                emptyHint="No training matrix rows for your company yet."
              />
              <div className={styles.tableFooter}>
                <p className={styles.muted}>
                  Showing 1 to {Math.min(5, p.matrixRows.length)} of{" "}
                  {p.matrixTotal} candidates
                </p>
                <button
                  type="button"
                  className={styles.linkBtn}
                  onClick={() => p.onNavigate("training-matrix")}
                >
                  View full training matrix →
                </button>
              </div>
            </>
          )}
        </section>

        <div className={styles.dashSide}>
          <section className={styles.card}>
            <div className={styles.cardHeader}>
              <h2 className={styles.cardTitle}>Customer Documents</h2>
              <button
                type="button"
                className={styles.linkBtn}
                onClick={() => p.onNavigate("documents")}
              >
                View all
              </button>
            </div>
            <div className={styles.docGrid}>
              {p.documentTiles.length === 0 ? (
                <p className={styles.muted}>No visible documents for your company.</p>
              ) : (
                p.documentTiles.map((doc) => (
                  <button
                    key={doc.id}
                    type="button"
                    className={`${styles.docTile} ${
                      doc.kind === "pdf" ? styles.pdfTile : ""
                    }`}
                    onClick={() => p.onNavigate("documents")}
                  >
                    {doc.kind === "pdf" ? <IconPdf /> : <IconFolder />}
                    <span className={styles.docLabel}>{doc.label}</span>
                    <span className={styles.docMeta}>{doc.meta}</span>
                  </button>
                ))
              )}
            </div>
          </section>

          <section className={styles.card}>
            <div className={styles.cardHeader}>
              <h2 className={styles.cardTitle}>NVQ Progress</h2>
              <button
                type="button"
                className={styles.linkBtn}
                onClick={() => p.onNavigate("nvq-progress")}
              >
                View all
              </button>
            </div>
            <div className={styles.nvqList}>
              {p.nvqRows.length === 0 ? (
                <p className={styles.muted}>No visible NVQ records.</p>
              ) : (
                p.nvqRows.map((row) => (
                  <div key={row.id} className={styles.nvqRow}>
                    <div className={styles.nvqTop}>
                      <div>
                        <p className={styles.nvqName}>{row.name}</p>
                        <p className={styles.nvqCourse}>{row.course}</p>
                      </div>
                      <span className={styles.nvqPct}>{row.pct}%</span>
                    </div>
                    <div className={styles.progressTrack}>
                      <div
                        className={styles.progressFill}
                        style={{ width: row.pct + "%" }}
                      />
                    </div>
                  </div>
                ))
              )}
            </div>
          </section>
        </div>
      </div>

      <div className={styles.dashBottom}>
        <section className={styles.card}>
          <div className={styles.cardHeader}>
            <h2 className={styles.cardTitle}>Upcoming Events</h2>
            <button
              type="button"
              className={styles.linkBtn}
              onClick={() => p.onNavigate("events")}
            >
              View all
            </button>
          </div>
          <div className={styles.eventList}>
            {p.eventRows.length === 0 ? (
              <p className={styles.muted}>No upcoming visible events.</p>
            ) : (
              p.eventRows.map((ev) => {
                const badge = parseEventDate(ev.dateRaw);
                return (
                  <button
                    key={ev.id}
                    type="button"
                    className={styles.eventCard}
                    onClick={() => p.onNavigate("events")}
                  >
                    <span className={styles.dateBadge}>
                      <span className={styles.dateMonth}>{badge.month}</span>
                      <span className={styles.dateDay}>{badge.day}</span>
                    </span>
                    <span>
                      <p className={styles.eventTitle}>{ev.title}</p>
                      <p className={styles.eventMeta}>{ev.when}</p>
                      <p className={styles.eventMeta}>
                        <IconPin /> {ev.where}
                      </p>
                    </span>
                  </button>
                );
              })
            )}
          </div>
        </section>

        <section className={styles.card}>
          <div className={styles.cardHeader}>
            <h2 className={styles.cardTitle}>Offers &amp; Promotions</h2>
            <button
              type="button"
              className={styles.linkBtn}
              onClick={() => p.onNavigate("offers")}
            >
              View all
            </button>
          </div>
          <div className={styles.offerGrid}>
            {p.offerCards.length === 0 ? (
              <p className={styles.muted}>No visible offers right now.</p>
            ) : (
              p.offerCards.map((offer) => (
                <button
                  key={offer.id}
                  type="button"
                  className={styles.offerCard}
                  onClick={() => p.onNavigate("offers")}
                >
                  <span className={styles.offerBadge}>{offer.badge}</span>
                  <p className={styles.offerTitle}>{offer.title}</p>
                  <span
                    className={
                      offer.code.indexOf("→") >= 0
                        ? styles.offerLink
                        : styles.offerCode
                    }
                  >
                    {offer.code}
                  </span>
                </button>
              ))
            )}
          </div>
        </section>
      </div>
    </div>
  );
};

function MatrixTable(props: {
  rows: PortalTableRow[];
  condensed: boolean;
  emptyHint: string;
}): React.ReactElement {
  const { rows, condensed, emptyHint } = props;
  if (rows.length === 0) {
    return <p className={styles.muted}>{emptyHint}</p>;
  }

  return (
    <div className={styles.tableWrap}>
      <table className={styles.table}>
        <thead>
          <tr>
            <th>Candidate</th>
            {!condensed && <th>Role</th>}
            {!condensed && <th>Key Training</th>}
            <th>Status</th>
            {!condensed && <th>Expiry Date</th>}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const name = row.cells[0] || "—";
            const status = row.cells[1] || "";
            const expiry = row.cells[2] || "—";
            const kind = statusKind(status);
            const pillClass =
              kind === "bad"
                ? styles.pillBad
                : kind === "warn"
                  ? styles.pillWarn
                  : styles.pillOk;
            return (
              <tr key={row.id}>
                <td>
                  <div className={styles.candidateCell}>
                    <span className={styles.avatarSm}>{initials(name)}</span>
                    {name}
                  </div>
                </td>
                {!condensed && <td className={styles.muted}>—</td>}
                {!condensed && <td className={styles.muted}>—</td>}
                <td>
                  {condensed ? (
                    <span className={`${styles.pill} ${pillClass}`} title={statusLabel(status)}>
                      {kind === "ok" ? "✓" : kind === "warn" ? "⏱" : "!"}
                    </span>
                  ) : (
                    <span className={`${styles.pill} ${pillClass}`}>
                      {statusLabel(status)}
                    </span>
                  )}
                </td>
                {!condensed && <td>{expiry}</td>}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function ListBody(props: {
  title: string;
  subtitle: string;
  stub?: string;
  headers: string[];
  rows: PortalTableRow[];
  loading: boolean;
  error: string | null;
  view: CustomerViewId;
  onNavigate: (v: CustomerViewId) => void;
}): React.ReactElement {
  const { title, subtitle, stub, headers, rows, loading, error, view } = props;

  if (stub) {
    return (
      <div className={styles.listPage}>
        <h2 className={styles.pageTitle}>{title}</h2>
        <p className={styles.pageSubtitle}>{subtitle}</p>
        <div className={styles.stubBox}>{stub}</div>
      </div>
    );
  }

  const isMatrix = view === "training-matrix";
  const isDocuments = view === "documents";

  return (
    <div className={styles.listPage}>
      <h2 className={styles.pageTitle}>{title}</h2>
      <p className={styles.pageSubtitle}>
        {subtitle}
        {!loading ? " · " + rows.length + " rows" : ""}
      </p>
      <div className={styles.card}>
        {loading && <p className={styles.muted}>Loading…</p>}
        {error && <p className={styles.error}>{error}</p>}
        {!loading && !error && isMatrix && (
          <MatrixTable
            rows={rows.slice(0, 50)}
            condensed={false}
            emptyHint="No matrix rows found for your company."
          />
        )}
        {!loading && !error && isDocuments && (
          <DocumentsTable rows={rows} />
        )}
        {!loading && !error && !isMatrix && !isDocuments && (
          <div className={styles.tableWrap}>
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
                    <td colSpan={headers.length || 1} className={styles.muted}>
                      No rows found.
                    </td>
                  </tr>
                ) : (
                  rows.map((row) => (
                    <tr key={row.id}>
                      {row.cells.map((c, i) => (
                        <td key={row.id + "-" + i}>{c}</td>
                      ))}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function DocumentsTable(props: { rows: PortalTableRow[] }): React.ReactElement {
  const { rows } = props;
  if (rows.length === 0) {
    return (
      <p className={styles.muted}>
        No documents have been shared with your account yet.
      </p>
    );
  }

  return (
    <div className={styles.tableWrap}>
      <table className={styles.table}>
        <thead>
          <tr>
            <th>Document Name</th>
            <th>Document Type</th>
            <th>Candidate Name</th>
            <th>Modified Date</th>
            <th>View</th>
            <th>Download</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const name = row.cells[0] || "—";
            const type = row.cells[1] || "—";
            const candidate = row.cells[2] || "—";
            const modified = row.cells[3] || "—";
            const viewUrl =
              (row.fields && (row.fields.__docViewUrl as string)) ||
              row.cells[4] ||
              "";
            const downloadUrl =
              (row.fields && (row.fields.__docDownloadUrl as string)) ||
              row.cells[5] ||
              "";
            const canDownload = Boolean(
              row.fields && row.fields.__docCanDownload
            );
            return (
              <tr key={row.id}>
                <td>{name}</td>
                <td>{type}</td>
                <td>{candidate}</td>
                <td>{modified}</td>
                <td>
                  {viewUrl ? (
                    <a
                      className={styles.linkBtn}
                      href={viewUrl}
                      target="_blank"
                      rel="noreferrer"
                    >
                      View
                    </a>
                  ) : (
                    <span className={styles.muted}>—</span>
                  )}
                </td>
                <td>
                  {canDownload && downloadUrl ? (
                    <a
                      className={styles.linkBtn}
                      href={downloadUrl}
                      target="_blank"
                      rel="noreferrer"
                      download={true}
                    >
                      Download
                    </a>
                  ) : (
                    <span className={styles.muted}>—</span>
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

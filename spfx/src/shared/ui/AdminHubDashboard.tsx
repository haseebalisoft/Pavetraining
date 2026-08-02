import * as React from "react";

import type { AdminViewId } from "./nav";
import styles from "./portal.module.scss";

export interface AdminHubDashboardProps {
  userEmail: string;
  counts: Record<string, number>;
  onNavigate: (id: AdminViewId) => void;
}

const ACTION_TILES: Array<{
  id: AdminViewId;
  title: string;
  hint: string;
  icon: string;
}> = [
  { id: "companies", title: "Companies", hint: "Company List", icon: "C" },
  { id: "workforce", title: "Workforce", hint: "Candidates", icon: "W" },
  { id: "documents", title: "Documents", hint: "Upload & visibility", icon: "D" },
  { id: "events", title: "Calendar", hint: "Bookings", icon: "B" },
  { id: "training-matrix", title: "Matrix", hint: "Training expiries", icon: "M" },
  { id: "training-records", title: "Registers", hint: "NPORS · EUSR · more", icon: "R" },
  { id: "offers", title: "Offers", hint: "Promotions", icon: "O" },
  { id: "permissions", title: "Permissions", hint: "Portal access", icon: "P" },
  { id: "bulk-upload", title: "Bulk upload", hint: "CSV import", icon: "U" },
];

const RESOURCE_TILES: Array<{
  id: AdminViewId;
  title: string;
  description: string;
  tone: "lime" | "charcoal" | "forest" | "slate" | "moss" | "ink";
}> = [
  {
    id: "training-matrix",
    title: "Training Matrix",
    description: "Wide expiry grid and sync",
    tone: "lime",
  },
  {
    id: "training-records",
    title: "Training Registers",
    description: "NPORS, EUSR, Streetworks, In-House",
    tone: "charcoal",
  },
  {
    id: "documents",
    title: "Customer Documents",
    description: "Folders, uploads, visibility",
    tone: "forest",
  },
  {
    id: "events",
    title: "Calendar / Bookings",
    description: "Events and Outlook sync",
    tone: "slate",
  },
  {
    id: "permissions",
    title: "Permissions",
    description: "Who can access the portals",
    tone: "moss",
  },
  {
    id: "logs",
    title: "Audit Log",
    description: "Portal activity history",
    tone: "ink",
  },
];

function displayNameFromEmail(email: string): string {
  const local = (email.split("@")[0] || email).trim();
  const parts = local.split(/[.\-_+\s]+/).filter(Boolean);
  if (parts.length === 0) return "Admin";
  return parts
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

const TONE_CLASS: Record<
  (typeof RESOURCE_TILES)[number]["tone"],
  string
> = {
  lime: styles.hubTone_lime,
  charcoal: styles.hubTone_charcoal,
  forest: styles.hubTone_forest,
  slate: styles.hubTone_slate,
  moss: styles.hubTone_moss,
  ink: styles.hubTone_ink,
};

function friendlyStatLabel(key: string): string {
  const map: Record<string, string> = {
    companies: "Active companies",
    workforce: "Active candidates",
    matrix: "Matrix rows",
    documents: "Documents",
    events: "Upcoming bookings",
    nvq: "NVQs",
    permissions: "Permissions",
  };
  return map[key] || key;
}

export const AdminHubDashboard: React.FC<AdminHubDashboardProps> = (props) => {
  const { userEmail, counts, onNavigate } = props;
  const welcomeName = displayNameFromEmail(userEmail);
  const countKeys = Object.keys(counts);

  return (
    <div className={styles.hubPage}>
      <section className={styles.hubHero} aria-label="Welcome">
        <div className={styles.hubHeroInner}>
          <p className={styles.hubHeroEyebrow}>PAVE Training · Admin</p>
          <h1 className={styles.hubHeroTitle}>Welcome, {welcomeName}</h1>
          <p className={styles.hubHeroSubtitle}>
            Operations hub for companies, workforce, matrix, bookings, and
            customer access — powered by SharePoint.
          </p>
          <div className={styles.hubHeroPanel}>
            <div>
              <p className={styles.hubHeroPanelLabel}>Operations</p>
              <p className={styles.hubHeroPanelHint}>
                Use the tiles below for everyday admin tasks. All lists open with
                full create, edit, delete, and export — same SharePoint data as
                the Next.js admin.
              </p>
            </div>
            <div>
              <p className={styles.hubHeroPanelLabel}>Signed in</p>
              <p className={styles.hubHeroPanelHint}>{userEmail}</p>
            </div>
          </div>
        </div>
      </section>

      <section className={styles.hubActionGrid} aria-label="Quick actions">
        {ACTION_TILES.map((tile) => (
          <button
            key={tile.id}
            type="button"
            className={styles.hubActionTile}
            onClick={() => onNavigate(tile.id)}
            aria-label={`Open ${tile.title}`}
          >
            <span className={styles.hubActionIcon} aria-hidden>
              {tile.icon}
            </span>
            <span className={styles.hubActionTitle}>{tile.title}</span>
            <span className={styles.hubActionHint}>{tile.hint}</span>
          </button>
        ))}
      </section>

      <section className={styles.hubResources} aria-label="Top resources">
        <div className={styles.hubResourcesHeader}>
          <h2 className={styles.hubResourcesTitle}>Top resources</h2>
          <p className={styles.hubResourcesSubtitle}>
            Jump into the lists you use most.
          </p>
        </div>
        <div className={styles.hubResourceGrid}>
          {RESOURCE_TILES.map((tile) => (
            <button
              key={tile.id}
              type="button"
              className={`${styles.hubResourceTile} ${TONE_CLASS[tile.tone]}`}
              onClick={() => onNavigate(tile.id)}
              aria-label={`Open ${tile.title}`}
            >
              <strong>{tile.title}</strong>
              <span>{tile.description}</span>
            </button>
          ))}
        </div>
      </section>

      <section className={styles.stats} aria-label="Dashboard statistics">
        {countKeys.length === 0 ? (
          <p className={styles.muted}>Loading SharePoint counts…</p>
        ) : (
          countKeys.map((key) => (
            <article key={key} className={styles.stat}>
              <p className={styles.statValue}>{counts[key]}</p>
              <p className={styles.statLabel}>{friendlyStatLabel(key)}</p>
            </article>
          ))
        )}
      </section>
    </div>
  );
};

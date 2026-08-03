"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useMemo, useState } from "react";

import { BrandLogo } from "@/components/brand/BrandLogo";
import styles from "./admin.module.css";

type NavLink = {
  type: "link";
  href: string;
  label: string;
  exact?: boolean;
};

type NavGroup = {
  type: "group";
  id: string;
  label: string;
  children: Array<{ href: string; label: string }>;
};

type NavItem = NavLink | NavGroup;

const NAV_ITEMS: NavItem[] = [
  { type: "link", href: "/admin", label: "Dashboard", exact: true },
  { type: "link", href: "/admin/companies", label: "Companies" },
  { type: "link", href: "/admin/departments", label: "Departments" },
  { type: "link", href: "/admin/workforce", label: "Workforce / Candidates" },
  { type: "link", href: "/admin/permissions", label: "Permissions & Access" },
  { type: "link", href: "/admin/training-matrix", label: "Training Matrix" },
  {
    type: "group",
    id: "registers",
    label: "Registers",
    children: [
      { href: "/admin/training-records/npors", label: "NPORS Register" },
      { href: "/admin/training-records/eusr", label: "EUSR Register" },
      {
        href: "/admin/training-records/streetworks",
        label: "Streetworks Training",
      },
      {
        href: "/admin/training-records/in-house",
        label: "In-House Certificates",
      },
      { href: "/admin/nvq", label: "NVQ Register" },
    ],
  },
  { type: "link", href: "/admin/documents", label: "Documents" },
  { type: "link", href: "/admin/events", label: "Calendar / Bookings" },
  { type: "link", href: "/admin/notifications", label: "Notifications" },
  { type: "link", href: "/admin/bulk-upload", label: "Bulk Upload" },
  { type: "link", href: "/admin/logs", label: "Audit / Activity Log" },
  { type: "link", href: "/admin/settings", label: "Settings" },
];

const REGISTER_PREFIXES = [
  "/admin/training-records",
  "/admin/nvq",
] as const;

interface AdminSidebarProps {
  email: string;
  signOutAction: () => Promise<void>;
}

function accountInitials(email: string): string {
  const local = email.split("@")[0]?.trim() || email.trim();
  const parts = local.split(/[.\-_+\s]+/).filter(Boolean);
  if (parts.length >= 2) {
    return `${parts[0]![0] ?? ""}${parts[1]![0] ?? ""}`.toUpperCase();
  }
  return local.slice(0, 2).toUpperCase() || "?";
}

function isLinkActive(
  pathname: string,
  href: string,
  exact?: boolean,
): boolean {
  if (exact) return pathname === href;
  return pathname === href || pathname.startsWith(`${href}/`);
}

function isRegistersPath(pathname: string): boolean {
  return REGISTER_PREFIXES.some(
    (prefix) => pathname === prefix || pathname.startsWith(`${prefix}/`),
  );
}

export function AdminSidebar({ email, signOutAction }: AdminSidebarProps) {
  const pathname = usePathname();
  const [open, setOpen] = useState(false);
  const registersActive = useMemo(() => isRegistersPath(pathname), [pathname]);
  const [registersOpen, setRegistersOpen] = useState(registersActive);

  useEffect(() => {
    setOpen(false);
  }, [pathname]);

  useEffect(() => {
    if (registersActive) {
      setRegistersOpen(true);
    }
  }, [registersActive]);

  useEffect(() => {
    if (!open) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [open]);

  return (
    <aside className={`${styles.sidebar} ${open ? styles.sidebarOpen : ""}`}>
      <div className={styles.sidebarTop}>
        <div className={styles.brandBlock}>
          <div className={styles.brandLogoPanel}>
            <BrandLogo variant="compact" priority />
            <p className={styles.brandTagline}>Paving the way in industry</p>
          </div>
        </div>
        <button
          type="button"
          className={styles.menuToggle}
          aria-expanded={open}
          aria-controls="admin-nav"
          onClick={() => setOpen((value) => !value)}
        >
          {open ? "Close" : "Menu"}
        </button>
      </div>

      <p className={styles.navSectionLabel}>PAVE Admin Portal</p>

      <nav
        id="admin-nav"
        className={`${styles.nav} ${open ? styles.navOpen : ""}`}
        aria-label="Admin"
      >
        {NAV_ITEMS.map((item) => {
          if (item.type === "link") {
            const active = isLinkActive(pathname, item.href, item.exact);
            return (
              <Link
                key={item.href}
                href={item.href}
                className={`${styles.navLink} ${active ? styles.navLinkActive : ""}`}
              >
                {item.label}
              </Link>
            );
          }

          return (
            <div key={item.id} className={styles.navGroup}>
              <button
                type="button"
                className={`${styles.navGroupToggle} ${registersActive ? styles.navGroupToggleActive : ""}`}
                aria-expanded={registersOpen}
                onClick={() => setRegistersOpen((value) => !value)}
              >
                <span>{item.label}</span>
                <span className={styles.navGroupChevron} aria-hidden="true">
                  {registersOpen ? "▾" : "▸"}
                </span>
              </button>
              {registersOpen ? (
                <div className={styles.navGroupChildren}>
                  {item.children.map((child) => {
                    const active = isLinkActive(pathname, child.href);
                    return (
                      <Link
                        key={child.href}
                        href={child.href}
                        className={`${styles.navChildLink} ${active ? styles.navLinkActive : ""}`}
                      >
                        {child.label}
                      </Link>
                    );
                  })}
                </div>
              ) : null}
            </div>
          );
        })}
      </nav>

      <div className={styles.sidebarFooter}>
        <div className={styles.accountRow}>
          <span className={styles.accountAvatar} aria-hidden="true">
            {accountInitials(email)}
          </span>
          <p className={styles.sidebarEmail} title={email}>
            {email}
          </p>
        </div>
        <form action={signOutAction}>
          <button className={styles.signOutButton} type="submit">
            Sign out
          </button>
        </form>
      </div>
    </aside>
  );
}

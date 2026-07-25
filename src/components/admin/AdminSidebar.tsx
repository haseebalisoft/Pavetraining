"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useState } from "react";

import { BrandLogo } from "@/components/brand/BrandLogo";
import styles from "./admin.module.css";

const NAV_ITEMS = [
  { href: "/admin", label: "Dashboard", exact: true },
  { href: "/admin/companies", label: "Companies" },
  { href: "/admin/workforce", label: "Workforce" },
  { href: "/admin/training-matrix", label: "Training Matrix" },
  { href: "/admin/training-records", label: "Training Records" },
  { href: "/admin/nvq", label: "NVQ" },
  { href: "/admin/documents", label: "Documents" },
  { href: "/admin/events", label: "Events" },
  { href: "/admin/offers", label: "Offers" },
  { href: "/admin/permissions", label: "Permissions" },
  { href: "/admin/automation", label: "Automation" },
  { href: "/admin/logs", label: "Logs" },
] as const;

interface AdminSidebarProps {
  email: string;
  signOutAction: () => Promise<void>;
}

export function AdminSidebar({ email, signOutAction }: AdminSidebarProps) {
  const pathname = usePathname();
  const [open, setOpen] = useState(false);

  useEffect(() => {
    setOpen(false);
  }, [pathname]);

  return (
    <aside className={`${styles.sidebar} ${open ? styles.sidebarOpen : ""}`}>
      <div className={styles.sidebarTop}>
        <div className={styles.brandBlock}>
          <div className={styles.brandLogoPanel}>
            <BrandLogo variant="compact" priority />
          </div>
          <p className={styles.brandMeta}>Admin operations</p>
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

      <nav
        id="admin-nav"
        className={`${styles.nav} ${open ? styles.navOpen : ""}`}
        aria-label="Admin"
      >
        {NAV_ITEMS.map((item) => {
          const exact = "exact" in item && item.exact;
          const active = exact
            ? pathname === item.href
            : pathname === item.href || pathname.startsWith(`${item.href}/`);

          return (
            <Link
              key={item.href}
              href={item.href}
              className={`${styles.navLink} ${active ? styles.navLinkActive : ""}`}
            >
              {item.label}
            </Link>
          );
        })}
      </nav>

      <div className={styles.sidebarFooter}>
        <p className={styles.sidebarEmail}>{email}</p>
        <form action={signOutAction}>
          <button className={styles.signOutButton} type="submit">
            Sign out
          </button>
        </form>
      </div>
    </aside>
  );
}

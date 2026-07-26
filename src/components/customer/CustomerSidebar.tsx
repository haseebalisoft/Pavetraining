"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useState } from "react";

import { BrandLogo } from "@/components/brand/BrandLogo";
import styles from "./customer.module.css";

const NAV_ITEMS = [
  { href: "/customer", label: "Training Matrix", exact: true },
  { href: "/customer/dashboard", label: "Dashboard" },
  { href: "/customer/candidates", label: "Candidates / Workforce" },
  { href: "/customer/training-records", label: "Training Records" },
  { href: "/customer/nvq-progress", label: "NVQ Progress" },
  { href: "/customer/documents", label: "Documents" },
  { href: "/customer/events", label: "Events / Bookings" },
  { href: "/customer/offers", label: "Offers" },
  { href: "/customer/support", label: "Support" },
] as const;

interface CustomerSidebarProps {
  email: string;
  companyName: string;
  roleLabel: string;
  accessLabel: string;
  canDownload: boolean;
  signOutAction: () => Promise<void>;
}

export function CustomerSidebar({
  email,
  companyName,
  roleLabel,
  accessLabel,
  canDownload,
  signOutAction,
}: CustomerSidebarProps) {
  const pathname = usePathname();
  const [open, setOpen] = useState(false);

  useEffect(() => {
    setOpen(false);
  }, [pathname]);

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
          </div>
          <p className={styles.brandMeta}>PAVE Training Customer Portal</p>
          <p className={styles.companyChip}>{companyName}</p>
          <p className={styles.accessChip}>Role: {roleLabel}</p>
          <p className={styles.accessChip}>Access: {accessLabel}</p>
          <p className={styles.accessChip}>
            Downloads: {canDownload ? "Enabled" : "Disabled"}
          </p>
        </div>
        <button
          type="button"
          className={styles.menuToggle}
          aria-expanded={open}
          aria-controls="customer-nav"
          onClick={() => setOpen((value) => !value)}
        >
          {open ? "Close" : "Menu"}
        </button>
      </div>

      <nav
        id="customer-nav"
        className={`${styles.nav} ${open ? styles.navOpen : ""}`}
        aria-label="Customer"
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

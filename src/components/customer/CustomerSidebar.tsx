"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useRef, useState, type MouseEvent } from "react";

import { BrandLogo } from "@/components/brand/BrandLogo";
import { SlideOverPanel } from "@/components/ui/SlideOverPanel";
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

/**
 * Prefer a real display name when one is passed later; today customer layout
 * only has email, so fall back to the local-part (e.g. john.smith → JS).
 */
function avatarInitials(email: string, displayName?: string | null): string {
  const fromName = displayName
    ?.trim()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => Array.from(part).find((ch) => /[a-z0-9]/i.test(ch)) ?? "")
    .join("");
  if (fromName) return fromName.toUpperCase();

  const localPart = email.split("@")[0]?.trim() ?? "";
  const segments = localPart.split(/[.\-_+\s]+/).filter(Boolean);
  if (segments.length >= 2) {
    const a = Array.from(segments[0]!).find((ch) => /[a-z0-9]/i.test(ch));
    const b = Array.from(segments[1]!).find((ch) => /[a-z0-9]/i.test(ch));
    if (a && b) return `${a}${b}`.toUpperCase();
  }
  const first = Array.from(localPart).find((ch) => /[a-z0-9]/i.test(ch));
  return (first ?? "A").toUpperCase();
}

function AccountIcon({
  type,
}: {
  type: "company" | "role" | "access" | "downloads";
}) {
  if (type === "company") {
    return <path d="M4 20V7l8-4 8 4v13M8 20v-5h8v5M8 9h.01M12 9h.01M16 9h.01" />;
  }
  if (type === "role") {
    return (
      <>
        <circle cx="12" cy="8" r="4" />
        <path d="M4.5 21a7.5 7.5 0 0 1 15 0" />
      </>
    );
  }
  if (type === "access") {
    return (
      <>
        <rect x="3" y="5" width="18" height="14" rx="2" />
        <path d="M8 10h8M8 14h5" />
      </>
    );
  }
  return (
    <>
      <path d="M12 3v12M7 10l5 5 5-5" />
      <path d="M5 21h14" />
    </>
  );
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
  const [navOpen, setNavOpen] = useState(false);
  const [accountOpen, setAccountOpen] = useState(false);
  const accountTriggerRef = useRef<HTMLButtonElement>(null);
  const initials = avatarInitials(email);

  useEffect(() => {
    if (!navOpen) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setNavOpen(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [navOpen]);

  function openAccountPanel(event: MouseEvent<HTMLButtonElement>) {
    event.preventDefault();
    event.stopPropagation();
    // Close mobile nav first so the sheet isn't fighting the collapsible menu.
    setNavOpen(false);
    setAccountOpen(true);
  }

  return (
    <aside className={`${styles.sidebar} ${navOpen ? styles.sidebarOpen : ""}`}>
      <div className={styles.sidebarTop}>
        <div className={styles.brandBlock}>
          <div className={styles.brandLogoPanel}>
            <BrandLogo variant="compact" priority />
          </div>
          <p className={styles.brandMeta}>PAVE Training Customer Portal</p>
          <button
            ref={accountTriggerRef}
            type="button"
            className={styles.accountTrigger}
            aria-haspopup="dialog"
            aria-expanded={accountOpen}
            onClick={openAccountPanel}
          >
            <span className={styles.accountAvatar} aria-hidden="true">
              {initials}
            </span>
            <span className={styles.accountCompany}>{companyName}</span>
            <span className={styles.accountChevron} aria-hidden="true">
              ›
            </span>
            <span className={styles.srOnly}>Open account details</span>
          </button>
        </div>
        <button
          type="button"
          className={styles.menuToggle}
          aria-expanded={navOpen}
          aria-controls="customer-nav"
          onClick={() => setNavOpen((value) => !value)}
        >
          {navOpen ? "Close" : "Menu"}
        </button>
      </div>

      <nav
        id="customer-nav"
        className={`${styles.nav} ${navOpen ? styles.navOpen : ""}`}
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
              onClick={() => setNavOpen(false)}
            >
              {item.label}
            </Link>
          );
        })}
      </nav>

      <SlideOverPanel
        open={accountOpen}
        title="Account details"
        onClose={() => setAccountOpen(false)}
        returnFocusRef={accountTriggerRef}
        header={
          <div className={styles.accountPanelIdentity}>
            <span className={styles.accountAvatarLarge} aria-hidden="true">
              {initials}
            </span>
            <p className={styles.accountPanelEmail}>{email}</p>
          </div>
        }
        footer={
          <form action={signOutAction} className={styles.accountSignOutForm}>
            <button className={styles.accountSignOutButton} type="submit">
              Sign out
            </button>
          </form>
        }
      >
        <div className={styles.accountInfoList}>
          {(
            [
              { type: "company" as const, label: "Company", value: companyName },
              { type: "role" as const, label: "Role", value: roleLabel },
              {
                type: "access" as const,
                label: "Access scope",
                value: accessLabel,
              },
            ] as const
          ).map((row) => (
            <div key={row.type} className={styles.accountInfoRow}>
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <AccountIcon type={row.type} />
              </svg>
              <div>
                <p>{row.label}</p>
                <strong>{row.value}</strong>
              </div>
            </div>
          ))}
          <div className={styles.accountInfoRow}>
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <AccountIcon type="downloads" />
            </svg>
            <div>
              <p>Downloads</p>
              <span
                className={
                  canDownload
                    ? styles.accountStatusEnabled
                    : styles.accountStatusDisabled
                }
              >
                {canDownload ? "Enabled" : "Disabled"}
              </span>
            </div>
          </div>
        </div>
      </SlideOverPanel>
    </aside>
  );
}

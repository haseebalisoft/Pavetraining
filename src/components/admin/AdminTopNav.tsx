"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useMemo, useRef, useState } from "react";

import { BrandLogo } from "@/components/brand/BrandLogo";
import styles from "./admin.module.css";

type TopLink = { href: string; label: string; exact?: boolean };

const PRIMARY_LINKS: TopLink[] = [
  { href: "/admin", label: "Home", exact: true },
  { href: "/admin/companies", label: "Companies" },
  { href: "/admin/workforce", label: "Workforce" },
  { href: "/admin/training-matrix", label: "Matrix" },
  { href: "/admin/documents", label: "Documents" },
  { href: "/admin/events", label: "Calendar" },
];

const REGISTER_LINKS: TopLink[] = [
  { href: "/admin/training-records", label: "All registers" },
  { href: "/admin/training-records/npors", label: "NPORS" },
  { href: "/admin/training-records/eusr", label: "EUSR" },
  { href: "/admin/training-records/streetworks", label: "Streetworks" },
  { href: "/admin/training-records/in-house", label: "In-House" },
  { href: "/admin/nvq", label: "NVQ" },
];

const MORE_LINKS: TopLink[] = [
  { href: "/admin/permissions", label: "Permissions" },
  { href: "/admin/offers", label: "Offers" },
  { href: "/admin/notifications", label: "Notifications" },
  { href: "/admin/bulk-upload", label: "Bulk Upload" },
  { href: "/admin/logs", label: "Audit Log" },
  { href: "/admin/settings", label: "Settings" },
];

const MOBILE_ALL_LINKS: TopLink[] = [
  ...PRIMARY_LINKS,
  ...REGISTER_LINKS,
  ...MORE_LINKS,
];

interface AdminTopNavProps {
  email: string;
  signOutAction: () => Promise<void>;
}

function isActive(pathname: string, href: string, exact?: boolean): boolean {
  if (exact) return pathname === href;
  return pathname === href || pathname.startsWith(`${href}/`);
}

function isRegistersPath(pathname: string): boolean {
  return (
    pathname === "/admin/training-records" ||
    pathname.startsWith("/admin/training-records/") ||
    pathname === "/admin/nvq" ||
    pathname.startsWith("/admin/nvq/")
  );
}

function isMorePath(pathname: string): boolean {
  return MORE_LINKS.some((link) => isActive(pathname, link.href));
}

function ChevronIcon() {
  return (
    <svg
      className={styles.topNavCaret}
      viewBox="0 0 16 16"
      aria-hidden="true"
      focusable="false"
    >
      <path
        d="m4 6 4 4 4-4"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

export function AdminTopNav({ email, signOutAction }: AdminTopNavProps) {
  const pathname = usePathname();
  const [mobileOpen, setMobileOpen] = useState(false);
  const [registersOpen, setRegistersOpen] = useState(false);
  const [moreOpen, setMoreOpen] = useState(false);
  const [profileOpen, setProfileOpen] = useState(false);
  const navRef = useRef<HTMLElement | null>(null);

  const registersActive = useMemo(() => isRegistersPath(pathname), [pathname]);
  const moreActive = useMemo(() => isMorePath(pathname), [pathname]);
  const profileInitial = email.trim().charAt(0).toUpperCase() || "A";

  useEffect(() => {
    setMobileOpen(false);
    setRegistersOpen(false);
    setMoreOpen(false);
    setProfileOpen(false);
  }, [pathname]);

  useEffect(() => {
    function onPointerDown(event: MouseEvent) {
      if (!navRef.current?.contains(event.target as Node)) {
        setRegistersOpen(false);
        setMoreOpen(false);
        setProfileOpen(false);
      }
    }
    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        setRegistersOpen(false);
        setMoreOpen(false);
        setProfileOpen(false);
        setMobileOpen(false);
      }
    }
    window.addEventListener("mousedown", onPointerDown);
    window.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener("mousedown", onPointerDown);
      window.removeEventListener("keydown", onKeyDown);
    };
  }, []);

  useEffect(() => {
    if (!mobileOpen) return;
    const previous = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = previous;
    };
  }, [mobileOpen]);

  return (
    <header className={styles.topNav} ref={navRef}>
      <div className={styles.topNavBar}>
        <div className={styles.topNavBrand}>
          <Link href="/admin" className={styles.topNavBrandLink}>
            <BrandLogo variant="mark" priority />
            <span className={styles.topNavBrandText}>PAVE HUB</span>
          </Link>
        </div>

        <nav className={styles.topNavLinksDesktop} aria-label="Admin">
          {PRIMARY_LINKS.map((link) => (
            <Link
              key={link.href}
              href={link.href}
              className={`${styles.topNavLink} ${
                isActive(pathname, link.href, link.exact)
                  ? styles.topNavLinkActive
                  : ""
              }`}
            >
              {link.label}
            </Link>
          ))}

          <div className={styles.topNavDropdown}>
            <button
              type="button"
              className={`${styles.topNavLink} ${styles.topNavDropdownToggle} ${
                registersActive ? styles.topNavLinkActive : ""
              }`}
              aria-expanded={registersOpen}
              onClick={() => {
                setRegistersOpen((value) => !value);
                setMoreOpen(false);
                setProfileOpen(false);
              }}
            >
              Registers
              <ChevronIcon />
            </button>
            {registersOpen ? (
              <div className={styles.topNavDropdownMenu} role="menu">
                {REGISTER_LINKS.map((link) => (
                  <Link
                    key={link.href}
                    href={link.href}
                    className={`${styles.topNavDropdownItem} ${
                      isActive(pathname, link.href)
                        ? styles.topNavDropdownItemActive
                        : ""
                    }`}
                    role="menuitem"
                  >
                    {link.label}
                  </Link>
                ))}
              </div>
            ) : null}
          </div>

          <div className={styles.topNavDropdown}>
            <button
              type="button"
              className={`${styles.topNavLink} ${styles.topNavDropdownToggle} ${
                moreActive ? styles.topNavLinkActive : ""
              }`}
              aria-expanded={moreOpen}
              onClick={() => {
                setMoreOpen((value) => !value);
                setRegistersOpen(false);
                setProfileOpen(false);
              }}
            >
              More
              <ChevronIcon />
            </button>
            {moreOpen ? (
              <div className={styles.topNavDropdownMenu} role="menu">
                {MORE_LINKS.map((link) => (
                  <Link
                    key={link.href}
                    href={link.href}
                    className={`${styles.topNavDropdownItem} ${
                      isActive(pathname, link.href)
                        ? styles.topNavDropdownItemActive
                        : ""
                    }`}
                    role="menuitem"
                  >
                    {link.label}
                  </Link>
                ))}
              </div>
            ) : null}
          </div>
        </nav>

        <div className={styles.topNavTrailing}>
          <div className={styles.topNavProfile}>
            <button
              type="button"
              className={`${styles.profileButton} ${
                profileOpen ? styles.profileButtonOpen : ""
              }`}
              aria-expanded={profileOpen}
              aria-haspopup="menu"
              aria-label="Account menu"
              title={email}
              onClick={() => {
                setProfileOpen((value) => !value);
                setRegistersOpen(false);
                setMoreOpen(false);
              }}
            >
              <span className={styles.profileAvatar} aria-hidden>
                {profileInitial}
              </span>
              <span className={styles.profileEmail}>{email}</span>
              <ChevronIcon />
            </button>
            {profileOpen ? (
              <div className={styles.profileMenu} role="menu">
                <div className={styles.profileMenuHeader}>
                  <div>
                    <p className={styles.profileMenuTitle}>Signed in</p>
                    <p className={styles.profileMenuEmail}>{email}</p>
                  </div>
                </div>
                <form action={signOutAction}>
                  <button
                    className={styles.profileSignOut}
                    type="submit"
                    role="menuitem"
                  >
                    Sign out
                  </button>
                </form>
              </div>
            ) : null}
          </div>

          <button
            type="button"
            className={styles.topNavMenuToggle}
            aria-expanded={mobileOpen}
            aria-controls="admin-mobile-drawer"
            onClick={() => {
              setMobileOpen((value) => !value);
              setProfileOpen(false);
            }}
          >
            <span className={styles.menuToggleBars} aria-hidden>
              <span />
              <span />
              <span />
            </span>
            <span>{mobileOpen ? "Close" : "Menu"}</span>
          </button>
        </div>
      </div>

      {mobileOpen ? (
        <div
          className={styles.mobileDrawerScrim}
          onClick={() => setMobileOpen(false)}
          aria-hidden
        />
      ) : null}

      <div
        id="admin-mobile-drawer"
        className={`${styles.mobileDrawer} ${
          mobileOpen ? styles.mobileDrawerOpen : ""
        }`}
        role="dialog"
        aria-modal="true"
        aria-label="Admin menu"
        hidden={!mobileOpen}
      >
        <div className={styles.mobileDrawerHeader}>
          <p className={styles.mobileDrawerTitle}>All options</p>
          <button
            type="button"
            className={styles.mobileDrawerClose}
            onClick={() => setMobileOpen(false)}
          >
            Close
          </button>
        </div>
        <nav className={styles.mobileDrawerNav} aria-label="Admin mobile">
          {MOBILE_ALL_LINKS.map((link) => (
            <Link
              key={`${link.href}-${link.label}`}
              href={link.href}
              className={`${styles.mobileDrawerLink} ${
                isActive(pathname, link.href, link.exact)
                  ? styles.mobileDrawerLinkActive
                  : ""
              }`}
            >
              {link.label}
            </Link>
          ))}
        </nav>
        <div className={styles.mobileDrawerProfile}>
          <div>
            <p className={styles.profileMenuTitle}>Signed in</p>
            <p className={styles.profileMenuEmail}>{email}</p>
          </div>
          <form action={signOutAction} className={styles.mobileDrawerSignOutForm}>
            <button className={styles.profileSignOut} type="submit">
              Sign out
            </button>
          </form>
        </div>
      </div>
    </header>
  );
}

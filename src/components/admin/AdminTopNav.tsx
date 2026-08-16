"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useMemo, useRef, useState } from "react";

import type { CustomerRoleType } from "@/types/models";

import { BrandLogo } from "@/components/brand/BrandLogo";
import styles from "./admin.module.css";
import {
  ADMIN_MORE_LINKS,
  ADMIN_PRIMARY_LINKS,
  ADMIN_REGISTER_LINKS,
  type AdminNavItem,
  canSeeAdminNavItem,
} from "./adminNavItems";

interface AdminTopNavProps {
  email: string;
  /** Human role label ("Admin" / "Training Manager") for the diagnostic panel. */
  roleLabel: string;
  /**
   * True iff this user should have full SharePoint-Admin power (literal SP
   * "Admin" OR hardcoded protected-admin email). Training Managers who only
   * have `canAccessAdmin` will get `false` and will not see SharePoint-admin-only
   * items like Bulk Upload.
   */
  isSharePointAdmin: boolean;
  /** Raw SharePoint RoleType value ("Admin" / "Training Manager"). Debug-only. */
  sharePointRoleType: string;
  /** Customer sub-role ("TrainingManager" / …); null for pure Admins. Debug-only. */
  customerRole: CustomerRoleType | null;
  /** True when the user is kept as admin by the hardcoded protected list. Debug-only. */
  isAlwaysAdminEmail: boolean;
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

function isMorePathFor(items: AdminNavItem[], pathname: string): boolean {
  return items.some((link) => isActive(pathname, link.href));
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

export function AdminTopNav({
  email,
  roleLabel,
  isSharePointAdmin,
  sharePointRoleType,
  customerRole,
  isAlwaysAdminEmail,
  signOutAction,
}: AdminTopNavProps) {
  const pathname = usePathname();
  const [mobileOpen, setMobileOpen] = useState(false);
  const [registersOpen, setRegistersOpen] = useState(false);
  const [moreOpen, setMoreOpen] = useState(false);
  const [profileOpen, setProfileOpen] = useState(false);
  const [debugOpen, setDebugOpen] = useState(false);
  const navRef = useRef<HTMLElement | null>(null);
  const ignoreScrimClickRef = useRef(false);

  const navContext = useMemo(
    () => ({ isSharePointAdmin }),
    [isSharePointAdmin],
  );
  const primaryLinks = useMemo(
    () => ADMIN_PRIMARY_LINKS.filter((item) => canSeeAdminNavItem(item, navContext)),
    [navContext],
  );
  const registerLinks = useMemo(
    () => ADMIN_REGISTER_LINKS.filter((item) => canSeeAdminNavItem(item, navContext)),
    [navContext],
  );
  const moreLinks = useMemo(
    () => ADMIN_MORE_LINKS.filter((item) => canSeeAdminNavItem(item, navContext)),
    [navContext],
  );
  const mobileAllLinks = useMemo(
    () => [...primaryLinks, ...registerLinks, ...moreLinks],
    [primaryLinks, registerLinks, moreLinks],
  );
  const allowedNavHrefs = useMemo(
    () => mobileAllLinks.map((item) => item.href),
    [mobileAllLinks],
  );

  const registersActive = useMemo(() => isRegistersPath(pathname), [pathname]);
  const moreActive = useMemo(
    () => isMorePathFor(moreLinks, pathname),
    [moreLinks, pathname],
  );
  const profileInitial = email.trim().charAt(0).toUpperCase() || "A";

  useEffect(() => {
    setMobileOpen(false);
    setRegistersOpen(false);
    setMoreOpen(false);
    setProfileOpen(false);
    setDebugOpen(false);
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
    // Mobile browsers synthesize a click on whatever appears under the finger
    // after touchend — usually the new full-screen scrim — which would close
    // the drawer instantly. Ignore scrim closes briefly after open.
    ignoreScrimClickRef.current = true;
    const release = window.setTimeout(() => {
      ignoreScrimClickRef.current = false;
    }, 400);
    const previous = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      window.clearTimeout(release);
      document.body.style.overflow = previous;
    };
  }, [mobileOpen]);

  function toggleMobileMenu() {
    setMobileOpen((value) => !value);
    setProfileOpen(false);
    setRegistersOpen(false);
    setMoreOpen(false);
  }

  function closeMobileMenu() {
    if (ignoreScrimClickRef.current) return;
    setMobileOpen(false);
  }

  return (
    <header
      className={`${styles.topNav} ${mobileOpen ? styles.topNavMenuOpen : ""}`}
      ref={navRef}
    >
      <div className={styles.topNavBar}>
        <div className={styles.topNavBrand}>
          <Link href="/admin" className={styles.topNavBrandLink}>
            <BrandLogo variant="mark" priority />
            <span className={styles.topNavBrandText}>PAVE HUB</span>
          </Link>
        </div>

        <nav className={styles.topNavLinksDesktop} aria-label="Admin">
          {primaryLinks.map((link) => (
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

          {registerLinks.length > 0 ? (
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
                  {registerLinks.map((link) => (
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
          ) : null}

          {moreLinks.length > 0 ? (
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
                  {moreLinks.map((link) => (
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
          ) : null}
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
                    <p className={styles.profileMenuRole}>
                      {roleLabel}
                      {isSharePointAdmin ? "" : " · limited admin"}
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  className={styles.profileDebugToggle}
                  aria-expanded={debugOpen}
                  onClick={() => setDebugOpen((value) => !value)}
                >
                  {debugOpen ? "Hide debug info" : "Show debug info"}
                </button>
                {debugOpen ? (
                  <dl className={styles.profileDebugList}>
                    <dt>Email</dt>
                    <dd>{email}</dd>
                    <dt>Resolved role</dt>
                    <dd>{roleLabel}</dd>
                    <dt>SharePoint RoleType</dt>
                    <dd>{sharePointRoleType || "—"}</dd>
                    <dt>Customer sub-role</dt>
                    <dd>{customerRole ?? "none (pure admin)"}</dd>
                    <dt>SharePoint Admin</dt>
                    <dd>
                      {isSharePointAdmin ? "yes" : "no"}
                      {isAlwaysAdminEmail
                        ? " · via hardcoded protected list"
                        : ""}
                    </dd>
                    <dt>Allowed nav items ({allowedNavHrefs.length})</dt>
                    <dd>
                      <ul className={styles.profileDebugItems}>
                        {allowedNavHrefs.map((href) => (
                          <li key={href}>{href}</li>
                        ))}
                      </ul>
                    </dd>
                  </dl>
                ) : null}
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
            onClick={toggleMobileMenu}
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
          onClick={closeMobileMenu}
          onTouchEnd={(event) => {
            // Swallow the ghost touch that would otherwise click the scrim.
            if (ignoreScrimClickRef.current) {
              event.preventDefault();
              event.stopPropagation();
            }
          }}
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
          {mobileAllLinks.map((link) => (
            <Link
              key={`${link.href}-${link.label}`}
              href={link.href}
              className={`${styles.mobileDrawerLink} ${
                isActive(pathname, link.href, link.exact)
                  ? styles.mobileDrawerLinkActive
                  : ""
              }`}
              onClick={() => setMobileOpen(false)}
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

"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type MouseEvent,
} from "react";

import { BrandLogo } from "@/components/brand/BrandLogo";
import { SlideOverPanel } from "@/components/ui/SlideOverPanel";
import styles from "./customer.module.css";

type TopLink = { href: string; label: string; exact?: boolean };

const PRIMARY_LINKS: TopLink[] = [
  { href: "/customer", label: "Training Matrix", exact: true },
  { href: "/customer/dashboard", label: "Dashboard" },
  { href: "/customer/courses", label: "Courses" },
  { href: "/customer/candidates", label: "Candidates" },
];

const TRAINING_LINKS: TopLink[] = [
  { href: "/customer/training-records", label: "Training Records" },
  { href: "/customer/nvq-progress", label: "NVQ Progress" },
  { href: "/customer/documents", label: "Documents" },
];

const MORE_LINKS: TopLink[] = [
  { href: "/customer/events", label: "Events / Bookings" },
  { href: "/customer/offers", label: "Offers" },
  { href: "/customer/support", label: "Support" },
];

const MOBILE_ALL_LINKS: TopLink[] = [
  ...PRIMARY_LINKS,
  ...TRAINING_LINKS,
  ...MORE_LINKS,
];

interface CustomerTopNavProps {
  email: string;
  companyName: string;
  roleLabel: string;
  accessLabel: string;
  canDownload: boolean;
  signOutAction: () => Promise<void>;
}

function isActive(pathname: string, href: string, exact?: boolean): boolean {
  if (exact) return pathname === href;
  return pathname === href || pathname.startsWith(`${href}/`);
}

function isTrainingPath(pathname: string): boolean {
  return TRAINING_LINKS.some((link) => isActive(pathname, link.href));
}

function isMorePath(pathname: string): boolean {
  return MORE_LINKS.some((link) => isActive(pathname, link.href));
}

function avatarInitials(email: string): string {
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

export function CustomerTopNav({
  email,
  companyName,
  roleLabel,
  accessLabel,
  canDownload,
  signOutAction,
}: CustomerTopNavProps) {
  const pathname = usePathname();
  const [mobileOpen, setMobileOpen] = useState(false);
  const [trainingOpen, setTrainingOpen] = useState(false);
  const [moreOpen, setMoreOpen] = useState(false);
  const [accountOpen, setAccountOpen] = useState(false);
  const navRef = useRef<HTMLElement | null>(null);
  const accountTriggerRef = useRef<HTMLButtonElement>(null);
  const ignoreScrimClickRef = useRef(false);
  const initials = avatarInitials(email);

  const trainingActive = useMemo(() => isTrainingPath(pathname), [pathname]);
  const moreActive = useMemo(() => isMorePath(pathname), [pathname]);

  useEffect(() => {
    setMobileOpen(false);
    setTrainingOpen(false);
    setMoreOpen(false);
  }, [pathname]);

  useEffect(() => {
    function onPointerDown(event: globalThis.MouseEvent) {
      if (!navRef.current?.contains(event.target as Node)) {
        setTrainingOpen(false);
        setMoreOpen(false);
      }
    }
    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        setTrainingOpen(false);
        setMoreOpen(false);
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
    setTrainingOpen(false);
    setMoreOpen(false);
  }

  function closeMobileMenu() {
    if (ignoreScrimClickRef.current) return;
    setMobileOpen(false);
  }

  function openAccountPanel(event: MouseEvent<HTMLButtonElement>) {
    event.preventDefault();
    event.stopPropagation();
    setMobileOpen(false);
    setTrainingOpen(false);
    setMoreOpen(false);
    setAccountOpen(true);
  }

  return (
    <header
      className={`${styles.topNav} ${mobileOpen ? styles.topNavMenuOpen : ""}`}
      ref={navRef}
    >
      <div className={styles.topNavBar}>
        <div className={styles.topNavBrand}>
          <Link href="/customer" className={styles.topNavBrandLink}>
            <BrandLogo variant="mark" priority />
            <span className={styles.topNavBrandText}>PAVE Training</span>
          </Link>
        </div>

        <nav className={styles.topNavLinksDesktop} aria-label="Customer">
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
                trainingActive ? styles.topNavLinkActive : ""
              }`}
              aria-expanded={trainingOpen}
              onClick={() => {
                setTrainingOpen((value) => !value);
                setMoreOpen(false);
              }}
            >
              Training
              <ChevronIcon />
            </button>
            {trainingOpen ? (
              <div className={styles.topNavDropdownMenu} role="menu">
                {TRAINING_LINKS.map((link) => (
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
                setTrainingOpen(false);
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
          <button
            ref={accountTriggerRef}
            type="button"
            className={styles.topNavProfileButton}
            aria-haspopup="dialog"
            aria-expanded={accountOpen}
            title={`${email} · ${companyName}`}
            onClick={openAccountPanel}
          >
            <span className={styles.topNavProfileAvatar} aria-hidden="true">
              {initials}
            </span>
            <span className={styles.topNavProfileCompany}>{companyName}</span>
            <ChevronIcon />
            <span className={styles.srOnly}>Open account details</span>
          </button>

          <button
            type="button"
            className={styles.topNavMenuToggle}
            aria-expanded={mobileOpen}
            aria-controls="customer-mobile-drawer"
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
            if (ignoreScrimClickRef.current) {
              event.preventDefault();
              event.stopPropagation();
            }
          }}
          aria-hidden
        />
      ) : null}

      <div
        id="customer-mobile-drawer"
        className={`${styles.mobileDrawer} ${
          mobileOpen ? styles.mobileDrawerOpen : ""
        }`}
        role="dialog"
        aria-modal="true"
        aria-label="Customer menu"
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
        <nav className={styles.mobileDrawerNav} aria-label="Customer mobile">
          {MOBILE_ALL_LINKS.map((link) => (
            <Link
              key={link.href}
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
          <button
            type="button"
            className={styles.mobileDrawerAccountButton}
            onClick={(event) => {
              openAccountPanel(event);
            }}
          >
            <span className={styles.topNavProfileAvatar} aria-hidden="true">
              {initials}
            </span>
            <span>
              <strong>{companyName}</strong>
              <em>{email}</em>
            </span>
          </button>
        </div>
      </div>

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
    </header>
  );
}

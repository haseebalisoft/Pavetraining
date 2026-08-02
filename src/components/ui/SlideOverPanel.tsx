"use client";

import {
  useEffect,
  useId,
  useRef,
  useState,
  type ReactNode,
  type RefObject,
} from "react";
import { createPortal } from "react-dom";

import styles from "./slideOverPanel.module.css";

const FOCUSABLE = [
  "a[href]",
  "button:not([disabled])",
  "input:not([disabled])",
  "select:not([disabled])",
  "textarea:not([disabled])",
  '[tabindex]:not([tabindex="-1"])',
].join(",");

interface SlideOverPanelProps {
  open: boolean;
  title: string;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  header?: ReactNode;
  wide?: boolean;
  returnFocusRef?: RefObject<HTMLElement | null>;
}

export function SlideOverPanel({
  open,
  title,
  onClose,
  children,
  footer,
  header,
  wide = false,
  returnFocusRef,
}: SlideOverPanelProps) {
  const titleId = useId();
  const panelRef = useRef<HTMLDivElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const onCloseRef = useRef(onClose);
  const ignoreBackdropClickRef = useRef(false);
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    onCloseRef.current = onClose;
  }, [onClose]);

  useEffect(() => {
    if (!open) return;

    // Opening from a click can synthesize a follow-up click on the new backdrop
    // (portal). Ignore backdrop closes for one frame so the sheet stays open.
    ignoreBackdropClickRef.current = true;
    const release = window.setTimeout(() => {
      ignoreBackdropClickRef.current = false;
    }, 320);

    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const returnFocusTarget =
      returnFocusRef?.current ?? previousFocusRef.current;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    const frame = window.requestAnimationFrame(() => panelRef.current?.focus());

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onCloseRef.current();
        return;
      }
      if (event.key !== "Tab" || !panelRef.current) return;

      const focusable = Array.from(
        panelRef.current.querySelectorAll<HTMLElement>(FOCUSABLE),
      );
      if (focusable.length === 0) {
        event.preventDefault();
        panelRef.current.focus();
        return;
      }

      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (
        event.shiftKey &&
        (document.activeElement === first ||
          document.activeElement === panelRef.current)
      ) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => {
      window.clearTimeout(release);
      window.cancelAnimationFrame(frame);
      window.removeEventListener("keydown", onKeyDown);
      document.body.style.overflow = previousOverflow;
      returnFocusTarget?.focus();
    };
  }, [open, returnFocusRef]);

  function handleBackdropClick() {
    if (ignoreBackdropClickRef.current) return;
    onClose();
  }

  if (!mounted || !open) return null;

  return createPortal(
    <div className={styles.root} data-slide-over-root="">
      <button
        type="button"
        className={styles.backdrop}
        aria-label={`Close ${title}`}
        onClick={handleBackdropClick}
      />
      <div
        ref={panelRef}
        className={`${styles.panel} ${wide ? styles.panelWide : ""}`}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        tabIndex={-1}
      >
        <div className={styles.header}>
          <div className={styles.headerContent}>
            {header ?? <h2 id={titleId}>{title}</h2>}
            {header ? (
              <span id={titleId} className={styles.srOnly}>
                {title}
              </span>
            ) : null}
          </div>
          <button
            type="button"
            className={styles.closeButton}
            aria-label={`Close ${title}`}
            onClick={onClose}
          >
            <span aria-hidden="true">×</span>
          </button>
        </div>
        <div className={styles.body}>{children}</div>
        {footer ? <div className={styles.footer}>{footer}</div> : null}
      </div>
    </div>,
    document.body,
  );
}

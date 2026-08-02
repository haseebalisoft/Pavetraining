"use client";

import { useState } from "react";

import styles from "./Thumbnail.module.css";

export type ThumbnailVariant = "company" | "person";

/**
 * Fixed-size image cell used in admin/list tables so rows keep a consistent
 * height whether or not an image exists. Renders the uploaded image cropped to
 * fill the box (object-fit: cover); when there is no image — or the image URL
 * fails to load — it falls back to a neutral placeholder icon inside the same
 * container, so empty and filled cells have identical visual weight.
 */
export function Thumbnail({
  src,
  alt = "",
  variant = "company",
  size = 40,
}: {
  src?: string | null;
  alt?: string;
  variant?: ThumbnailVariant;
  size?: number;
}) {
  const [failed, setFailed] = useState(false);
  // Reset the error state when the source changes (e.g. after an upload)
  // using React's "adjust state during render" pattern — no effect needed.
  const [seenSrc, setSeenSrc] = useState(src);
  if (src !== seenSrc) {
    setSeenSrc(src);
    setFailed(false);
  }

  const showImage = Boolean(src) && !failed;

  return (
    <span
      className={styles.thumb}
      style={{ width: size, height: size }}
      role="img"
      aria-label={alt || undefined}
    >
      {showImage ? (
        // eslint-disable-next-line @next/next/no-img-element
        <img
          className={styles.thumbImg}
          src={src as string}
          alt={alt}
          onError={() => setFailed(true)}
        />
      ) : (
        <PlaceholderIcon variant={variant} />
      )}
    </span>
  );
}

function PlaceholderIcon({ variant }: { variant: ThumbnailVariant }) {
  if (variant === "person") {
    return (
      <svg
        className={styles.thumbIcon}
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth={1.6}
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
        focusable="false"
      >
        <circle cx="12" cy="8" r="3.25" />
        <path d="M5.5 20a6.5 6.5 0 0 1 13 0" />
      </svg>
    );
  }

  return (
    <svg
      className={styles.thumbIcon}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.6}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
    >
      <path d="M3 21h18" />
      <path d="M6 21V5a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v16" />
      <path d="M14 21V10h4a1 1 0 0 1 1 1v10" />
      <path d="M9 8h1M9 12h1M9 16h1" />
    </svg>
  );
}

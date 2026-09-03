"use client";

import { useRef, useState, type ReactNode } from "react";

import { useAdminToast } from "@/components/admin/AdminToast";
import { readPublicApiError } from "@/lib/errors/publicMessages";

import styles from "./admin.module.css";

export function ImageUploadButton({
  uploadUrl,
  label = "Upload image",
  onUploaded,
  variant = "link",
  children,
}: {
  uploadUrl: string;
  label?: string;
  onUploaded?: () => Promise<void> | void;
  /** `link` = text control (tables). `button` = solid control (profiles). */
  variant?: "link" | "button";
  /** Optional trigger (e.g. clickable photo). Still shows the text/button label. */
  children?: ReactNode;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const { pushToast } = useAdminToast();
  const [uploading, setUploading] = useState(false);

  async function onFile(file: File | null) {
    if (!file) return;
    setUploading(true);
    try {
      const body = new FormData();
      body.set("file", file);
      const response = await fetch(uploadUrl, { method: "POST", body });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      pushToast(`${label} saved.`, "success");
      await onUploaded?.();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Upload failed.",
        "error",
      );
    } finally {
      setUploading(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  function openPicker() {
    if (uploading) return;
    inputRef.current?.click();
  }

  return (
    <span
      className={
        children || variant === "button"
          ? `${styles.imageUploadRow} ${styles.imageUploadRowStacked}`
          : styles.imageUploadRow
      }
    >
      <input
        ref={inputRef}
        type="file"
        accept="image/png,image/jpeg,image/webp,image/gif"
        hidden
        onChange={(event) => void onFile(event.target.files?.[0] ?? null)}
      />
      {children ? (
        <button
          type="button"
          className={styles.imageUploadHit}
          disabled={uploading}
          aria-label={uploading ? "Uploading…" : label}
          onClick={openPicker}
        >
          {children}
        </button>
      ) : null}
      <button
        type="button"
        className={
          variant === "button" ? styles.imageUploadButton : styles.linkButton
        }
        disabled={uploading}
        onClick={openPicker}
      >
        {uploading ? "Uploading…" : label}
      </button>
    </span>
  );
}

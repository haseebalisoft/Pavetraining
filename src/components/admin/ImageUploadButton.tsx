"use client";

import { useRef, useState } from "react";

import { useAdminToast } from "@/components/admin/AdminToast";
import { readPublicApiError } from "@/lib/errors/publicMessages";

import styles from "./admin.module.css";

export function ImageUploadButton({
  uploadUrl,
  label = "Upload image",
  currentUrl,
  onUploaded,
}: {
  uploadUrl: string;
  label?: string;
  currentUrl?: string | null;
  onUploaded?: () => Promise<void> | void;
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

  return (
    <span className={styles.imageUploadRow}>
      {currentUrl ? (
        // eslint-disable-next-line @next/next/no-img-element
        <img
          src={currentUrl}
          alt=""
          className={styles.imageUploadPreview}
        />
      ) : null}
      <input
        ref={inputRef}
        type="file"
        accept="image/png,image/jpeg,image/webp,image/gif"
        hidden
        onChange={(event) => void onFile(event.target.files?.[0] ?? null)}
      />
      <button
        type="button"
        className={styles.linkButton}
        disabled={uploading}
        onClick={() => inputRef.current?.click()}
      >
        {uploading ? "Uploading…" : label}
      </button>
    </span>
  );
}

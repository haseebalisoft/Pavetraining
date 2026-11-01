"use client";

import { useEffect, useState } from "react";

import { useAdminToast } from "@/components/admin/AdminToast";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { readPublicApiError } from "@/lib/errors/publicMessages";

import styles from "@/components/admin/admin.module.css";

interface EventDocument {
  id: string;
  name: string;
  customerVisible: boolean;
}

async function fetchEventDocuments(
  eventId: string,
): Promise<{ documents: EventDocument[]; error: string | null }> {
  try {
    const response = await fetch(
      `/api/admin/events/${encodeURIComponent(eventId)}/documents`,
      { cache: "no-store", signal: AbortSignal.timeout(60_000) },
    );
    if (!response.ok) throw new Error(await readPublicApiError(response));
    const payload = (await response.json()) as { documents?: EventDocument[] };
    return { documents: payload.documents ?? [], error: null };
  } catch (err) {
    return {
      documents: [],
      error: err instanceof Error ? err.message : "Could not load documents.",
    };
  }
}

/**
 * Attach / open / remove documents for one event. Files are stored with the
 * existing document storage (Customer Documents → {company}/Company
 * Documents/Events/{event id}); customers of that company open them from the
 * event when "Visible to customer" is ticked.
 */
export function AdminEventDocuments({ eventId }: { eventId: string }) {
  const { pushToast } = useAdminToast();
  const [documents, setDocuments] = useState<EventDocument[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [customerVisible, setCustomerVisible] = useState(true);

  /** Bumped after an upload / removal to reload the list. */
  const [revision, setRevision] = useState(0);
  const load = () => setRevision((value) => value + 1);

  useEffect(() => {
    let cancelled = false;
    fetchEventDocuments(eventId).then((result) => {
      if (cancelled) return;
      setDocuments(result.documents);
      setError(result.error);
    });
    return () => {
      cancelled = true;
    };
  }, [eventId, revision]);

  async function attach(files: FileList | null) {
    if (!files?.length) return;
    setBusy(true);
    let attached = 0;
    for (const file of Array.from(files)) {
      try {
        const form = new FormData();
        form.set("file", file);
        form.set("customerVisible", customerVisible ? "true" : "false");
        const response = await fetch(
          `/api/admin/events/${encodeURIComponent(eventId)}/documents`,
          { method: "POST", body: form, signal: AbortSignal.timeout(120_000) },
        );
        if (!response.ok) throw new Error(await readPublicApiError(response));
        attached += 1;
      } catch (err) {
        pushToast(
          `${file.name}: ${err instanceof Error ? err.message : "upload failed"}`,
          "error",
        );
      }
    }
    if (attached) {
      pushToast(`${attached} document(s) attached to the event.`, "success");
    }
    setBusy(false);
    load();
  }

  async function remove(doc: EventDocument) {
    if (!window.confirm(`Remove "${doc.name}" from this event? The file is deleted.`)) {
      return;
    }
    setBusy(true);
    try {
      const response = await fetch(`/api/admin/documents/${encodeURIComponent(doc.id)}`, {
        method: "DELETE",
        signal: AbortSignal.timeout(60_000),
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      pushToast(`Removed "${doc.name}".`, "success");
      load();
    } catch (err) {
      pushToast(err instanceof Error ? err.message : "Could not remove the document.", "error");
    } finally {
      setBusy(false);
    }
  }

  return (
    <section className={styles.eventDocuments} aria-label="Event documents">
      <h3 className={styles.fieldLabel}>Documents</h3>
      {error ? <p className={styles.formError}>{error}</p> : null}
      {documents === null ? (
        <p className={styles.bulkHint}>Loading documents…</p>
      ) : documents.length === 0 ? (
        <p className={styles.bulkHint}>No documents attached yet.</p>
      ) : (
        <ul className={styles.eventDocumentList}>
          {documents.map((doc) => (
            <li key={doc.id}>
              <a
                className={styles.link}
                href={`/api/admin/documents/${encodeURIComponent(doc.id)}/view`}
                target="_blank"
                rel="noreferrer"
              >
                {doc.name}
              </a>{" "}
              <StatusBadge
                label={doc.customerVisible ? "Visible to customer" : "Hidden"}
                tone={doc.customerVisible ? "ok" : "neutral"}
              />{" "}
              <a
                className={styles.linkButton}
                href={`/api/admin/documents/${encodeURIComponent(doc.id)}/download`}
              >
                Download
              </a>{" "}
              <button
                type="button"
                className={styles.linkButton}
                disabled={busy}
                onClick={() => void remove(doc)}
              >
                Remove
              </button>
            </li>
          ))}
        </ul>
      )}
      <div className={styles.eventDocumentActions}>
        <label className={styles.secondaryButton}>
          {busy ? "Uploading…" : "Attach document"}
          <input
            type="file"
            multiple
            className={styles.bulkFileInput}
            disabled={busy}
            onChange={(event) => {
              void attach(event.target.files);
              event.target.value = "";
            }}
          />
        </label>
        <label className={styles.bulkCheckRow}>
          <input
            type="checkbox"
            checked={customerVisible}
            onChange={(event) => setCustomerVisible(event.target.checked)}
          />
          <span>Visible to the event&apos;s company (customers can open and download)</span>
        </label>
      </div>
      <p className={styles.bulkHint}>
        Stored in Customer Documents under the event&apos;s company (Company
        Documents/Events). Up to 50 MB per file.
      </p>
    </section>
  );
}

"use client";

import { SlideOverPanel } from "@/components/ui/SlideOverPanel";
import { formatDate, formatTime } from "@/lib/utils/formatDate";
import { formatEventDuration } from "@/lib/utils/eventDuration";
import type { CustomerEventRecord } from "@/types/models";

import styles from "./customer.module.css";

function formatTimeRange(row: CustomerEventRecord): string {
  const start = formatTime(row.eventDate);
  const end = formatTime(row.endDate);
  if (start && end) return `${start} – ${end}`;
  return start ?? end ?? "Time TBC";
}

function locationLabel(row: CustomerEventRecord): string {
  const location = row.location?.trim();
  const address = row.trainingAddress?.trim();
  if (location && address && location.toLowerCase() !== address.toLowerCase()) {
    return `${location} · ${address}`;
  }
  return location || address || "—";
}

/**
 * Files an admin attached to this event (Events → Documents). Previously this
 * guessed by matching words in the event title against document names, which
 * could list unrelated files (e.g. any certificate containing "NPORS").
 */
export function documentsForEvent(
  event: CustomerEventRecord,
): Array<{ id: string; name: string; viewPath: string; downloadPath: string }> {
  return (event.documents ?? []).map((doc) => ({
    id: doc.id,
    name: doc.name,
    viewPath: `/api/customer/documents/${encodeURIComponent(doc.id)}/view`,
    downloadPath: `/api/customer/documents/${encodeURIComponent(doc.id)}/download`,
  }));
}

export function CustomerEventDetailPanel({
  event,
  onClose,
}: {
  event: CustomerEventRecord | null;
  onClose: () => void;
}) {
  const related = event ? documentsForEvent(event) : [];

  return (
    <SlideOverPanel
      open={event !== null}
      title={event?.title ?? "Event"}
      onClose={onClose}
    >
      {event ? (
        <div className={styles.eventDetail}>
          <dl className={styles.eventMetaList}>
            <div>
              <dt>Event name</dt>
              <dd>{event.title}</dd>
            </div>
            <div>
              <dt>Date</dt>
              <dd>
                {event.eventDate ? formatDate(event.eventDate) : "Date TBC"}
              </dd>
            </div>
            <div>
              <dt>Start / end time</dt>
              <dd>{formatTimeRange(event)}</dd>
            </div>
            <div>
              <dt>Duration</dt>
              <dd>
                {formatEventDuration(event.eventDate, event.endDate) ??
                  "Duration not set"}
              </dd>
            </div>
            <div className={styles.eventDetailWide}>
              <dt>Location</dt>
              <dd>{locationLabel(event)}</dd>
            </div>
          </dl>

          <section className={styles.eventDetailBlock} aria-label="Description">
            <h3>Description</h3>
            <p>
              {event.description?.trim() || "No description provided."}
            </p>
          </section>

          <section className={styles.eventDetailBlock} aria-label="Documents">
            <h3>Documents</h3>
            {related.length === 0 ? (
              <p className={styles.muted}>No documents for this event.</p>
            ) : (
              <ul className={styles.eventDocumentList}>
                {related.map((doc) => (
                  <li key={doc.id}>
                    <a
                      className={styles.link}
                      href={doc.viewPath}
                      target="_blank"
                      rel="noreferrer"
                    >
                      {doc.name}
                    </a>
                    <span className={styles.muted}> · </span>
                    <a className={styles.link} href={doc.downloadPath}>
                      Download
                    </a>
                  </li>
                ))}
              </ul>
            )}
          </section>
        </div>
      ) : null}
    </SlideOverPanel>
  );
}

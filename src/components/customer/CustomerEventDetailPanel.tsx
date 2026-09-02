"use client";

import { SlideOverPanel } from "@/components/ui/SlideOverPanel";
import { formatDate, formatTime } from "@/lib/utils/formatDate";
import { formatEventDuration } from "@/lib/utils/eventDuration";
import type {
  CustomerDocumentRecord,
  CustomerEventRecord,
} from "@/types/models";

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

export function documentsForEvent(
  event: CustomerEventRecord,
  documents: CustomerDocumentRecord[],
): CustomerDocumentRecord[] {
  const title = event.title.trim().toLowerCase();
  if (!title) return [];
  const tokens = title.split(/\s+/).filter((token) => token.length > 3);
  return documents.filter((doc) => {
    const haystack = `${doc.name} ${doc.documentType ?? ""}`.toLowerCase();
    if (haystack.includes(title)) return true;
    return tokens.some((token) => haystack.includes(token));
  });
}

export function CustomerEventDetailPanel({
  event,
  documents = [],
  onClose,
}: {
  event: CustomerEventRecord | null;
  documents?: CustomerDocumentRecord[];
  onClose: () => void;
}) {
  const related = event ? documentsForEvent(event, documents) : [];

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
                    {doc.viewPath ? (
                      <a
                        className={styles.link}
                        href={doc.viewPath}
                        target="_blank"
                        rel="noreferrer"
                      >
                        {doc.name}
                      </a>
                    ) : (
                      <span>{doc.name}</span>
                    )}
                    {doc.documentType?.trim() ? (
                      <span className={styles.muted}>
                        {" "}
                        · {doc.documentType}
                      </span>
                    ) : null}
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

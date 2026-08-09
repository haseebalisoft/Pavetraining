import Link from "next/link";

import {
  formatDate,
  formatDay,
  formatShortMonth,
  formatTime,
} from "@/lib/utils/formatDate";
import { formatEventDuration } from "@/lib/utils/eventDuration";
import type { CustomerEventRecord } from "@/types/models";

import styles from "./customerDashboard.module.css";

function PinIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 21s6-5.2 6-11a6 6 0 1 0-12 0c0 5.8 6 11 6 11Z" />
      <circle cx="12" cy="10" r="2" />
    </svg>
  );
}

function ClockIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <circle cx="12" cy="12" r="9" />
      <path d="M12 7v5l3 2" />
    </svg>
  );
}

function timeRange(event: CustomerEventRecord): string {
  const start = formatTime(event.eventDate);
  const end = formatTime(event.endDate);
  if (start && end) return `${start} – ${end}`;
  return start ?? end ?? "Time TBC";
}

function durationLabel(event: CustomerEventRecord): string {
  return (
    formatEventDuration(event.eventDate, event.endDate) ?? "Duration not set"
  );
}

function eventStatusLabel(event: CustomerEventRecord): string {
  if (event.eventDate) return "Confirmed";
  return "Date TBC";
}

export function CustomerUpcomingEvents({
  events,
}: {
  events: CustomerEventRecord[];
}) {
  return (
    <section className={styles.eventsSection} aria-labelledby="upcoming-events">
      <div className={styles.sectionHeader}>
        <h2 id="upcoming-events">Upcoming events</h2>
        <Link href="/customer/events" className={styles.viewAllLink}>
          View all <span aria-hidden="true">→</span>
        </Link>
      </div>

      {events.length === 0 ? (
        <div className={styles.eventsEmpty}>
          <h3>No upcoming events</h3>
          <p>New customer-visible training events will appear here.</p>
        </div>
      ) : (
        <div className={styles.eventsGrid}>
          {events.slice(0, 3).map((event) => {
            const location =
              event.location?.trim() ||
              event.trainingAddress?.trim() ||
              "Location TBC";
            const status = eventStatusLabel(event);
            return (
              <article key={event.id} className={styles.eventCard}>
                <div
                  className={styles.eventDateStrip}
                  aria-label={formatDate(event.eventDate, "Date TBC")}
                >
                  <span>{formatDay(event.eventDate)}</span>
                  <strong>{formatShortMonth(event.eventDate)}</strong>
                </div>
                <div className={styles.eventCardBody}>
                  <h3>{event.title}</h3>
                  <p className={`${styles.eventMeta} ${styles.eventMetaPin}`}>
                    <PinIcon />
                    <span>{location}</span>
                  </p>
                  <p className={`${styles.eventMeta} ${styles.eventMetaTime}`}>
                    <ClockIcon />
                    <span>{timeRange(event)}</span>
                  </p>
                  <p className={`${styles.eventMeta} ${styles.eventMetaTime}`}>
                    <span>Duration: {durationLabel(event)}</span>
                  </p>
                  <span
                    className={
                      status === "Confirmed"
                        ? styles.eventStatusOk
                        : styles.eventStatusNeutral
                    }
                  >
                    {status}
                  </span>
                </div>
              </article>
            );
          })}
        </div>
      )}
    </section>
  );
}

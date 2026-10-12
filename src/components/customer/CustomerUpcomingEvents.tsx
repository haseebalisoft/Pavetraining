import Link from "next/link";

import {
  formatDate,
  formatDay,
  formatShortMonth,
  formatTime,
} from "@/lib/utils/formatDate";
import type { CustomerEventRecord } from "@/types/models";

import { DashboardIcon } from "./DashboardIcons";
import styles from "./customerDashboard.module.css";

function timeRange(event: CustomerEventRecord): string {
  const start = formatTime(event.eventDate);
  const end = formatTime(event.endDate);
  if (start && end) return `${start}–${end}`;
  return start ?? end ?? "Time TBC";
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
    <section className={styles.splitColumn} aria-labelledby="upcoming-events">
      <div className={styles.sectionHeader}>
        <h2 id="upcoming-events">Upcoming events</h2>
        <Link href="/customer/events" className={styles.viewAllLink}>
          View all <DashboardIcon name="arrow" />
        </Link>
      </div>

      <div className={styles.listCard}>
        {events.length === 0 ? (
          <div className={styles.listEmpty}>
            <strong>No upcoming events</strong>
            <span>New customer-visible training events will appear here.</span>
          </div>
        ) : (
          events.slice(0, 3).map((event) => {
            const location =
              event.location?.trim() ||
              event.trainingAddress?.trim() ||
              "Location TBC";
            const status = eventStatusLabel(event);
            return (
              <Link
                key={event.id}
                href={`/customer/events?event=${encodeURIComponent(event.id)}`}
                className={styles.eventRow}
              >
                <span
                  className={styles.eventDateBlock}
                  aria-label={formatDate(event.eventDate, "Date TBC")}
                >
                  <strong>{formatDay(event.eventDate)}</strong>
                  <span>{formatShortMonth(event.eventDate)}</span>
                </span>
                <span className={styles.eventRowBody}>
                  <span className={styles.eventRowTitle}>{event.title}</span>
                  <span className={styles.eventRowMeta}>
                    <DashboardIcon name="pin" />
                    {location}
                  </span>
                  <span className={styles.eventRowMeta}>
                    <DashboardIcon name="clock" />
                    {timeRange(event)}
                  </span>
                </span>
                <span
                  className={
                    status === "Confirmed"
                      ? styles.statusPillActive
                      : styles.statusPillPending
                  }
                >
                  {status}
                </span>
                <DashboardIcon name="chevron" className={styles.rowChevron} />
              </Link>
            );
          })
        )}
      </div>
    </section>
  );
}

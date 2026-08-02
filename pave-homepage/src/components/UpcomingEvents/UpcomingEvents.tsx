import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import type { CalendarEvent } from "../../models";
import { EventsService } from "../../services/EventsService";
import { paveTheme } from "../../theme/paveTheme";
import styles from "./UpcomingEvents.module.scss";

export interface IUpcomingEventsProps {
  context: WebPartContext;
  eventsAddUrl: string;
}

function monthAbbrev(iso: string | null): string {
  if (!iso) return "—";
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleString("en-GB", { month: "short" }).toUpperCase();
}

function dayNumber(iso: string | null): string {
  if (!iso) return "–";
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "–";
  return String(date.getDate());
}

export const UpcomingEvents: React.FC<IUpcomingEventsProps> = (props) => {
  const { context, eventsAddUrl } = props;
  const [events, setEvents] = React.useState<CalendarEvent[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [failed, setFailed] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      setFailed(false);
      try {
        const rows = await EventsService.getUpcoming(context, 5);
        if (!cancelled) {
          setEvents(rows.slice(0, 5));
          setFailed(false);
        }
      } catch (error) {
        console.warn("[UpcomingEvents] Unable to load events", error);
        if (!cancelled) {
          setEvents([]);
          setFailed(true);
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [context]);

  const addUrl = (eventsAddUrl || "").trim();

  return (
    <section
      className={styles.section}
      aria-label="Upcoming events"
      style={
        {
          ["--pave-green" as string]: paveTheme.green,
          ["--pave-charcoal" as string]: paveTheme.charcoal,
          ["--pave-charcoal-dark" as string]: paveTheme.charcoalDark,
          ["--pave-font-body" as string]: paveTheme.fontBody,
        } as React.CSSProperties
      }
    >
      <div className={styles.header}>
        <h2 className={styles.title}>Upcoming events</h2>
        {addUrl ? (
          <a
            className={styles.addLink}
            href={addUrl}
            aria-label="Add a new event"
          >
            Add
          </a>
        ) : (
          <span className={styles.addMuted}>Add</span>
        )}
      </div>

      {loading ? (
        <ul className={styles.list} aria-busy="true" aria-label="Loading events">
          {[0, 1, 2].map((key) => (
            <li key={key} className={`${styles.row} ${styles.skeleton}`} />
          ))}
        </ul>
      ) : failed ? (
        <p className={styles.loadNote}>Unable to load — check permissions</p>
      ) : events.length === 0 ? (
        <p className={styles.empty}>No upcoming events.</p>
      ) : (
        <ul className={styles.list}>
          {events.map((event) => (
            <li key={event.id} className={styles.row}>
              <div className={styles.dateChip} aria-hidden="true">
                <span className={styles.month}>{monthAbbrev(event.start)}</span>
                <span className={styles.day}>{dayNumber(event.start)}</span>
              </div>
              <div className={styles.meta}>
                <p className={styles.eventTitle}>
                  {event.title || "Untitled event"}
                </p>
                {event.location ? (
                  <p className={styles.location}>{event.location}</p>
                ) : null}
              </div>
            </li>
          ))}
        </ul>
      )}
    </section>
  );
};

export default UpcomingEvents;

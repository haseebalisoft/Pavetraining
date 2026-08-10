"use client";

import { useMemo, useState } from "react";

import {
  CalendarGrid,
  type CalendarGridEvent,
  type CalendarView,
} from "@/components/calendar/CalendarGrid";
import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import { formatDate, formatDuration, formatTime } from "@/lib/utils/formatDate";
import type { CustomerEventRecord } from "@/types/models";

import styles from "./customer.module.css";
import sectionStyles from "./portalSections.module.css";

interface Props {
  companyName: string;
  records: CustomerEventRecord[];
}

type EventsViewMode = "list" | CalendarView;

function formatTimeRange(row: CustomerEventRecord): string {
  const start = formatTime(row.eventDate);
  const end = formatTime(row.endDate);
  if (start && end) return `${start} – ${end}`;
  return start ?? end ?? "Time TBC";
}

export function EventsView({ companyName, records }: Props) {
  const [search, setSearch] = useState("");
  const [view, setView] = useState<EventsViewMode>("month");
  const [cursor, setCursor] = useState(() => new Date());
  const [selectedId, setSelectedId] = useState<string | null>(null);

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) return records;
    return records.filter((row) =>
      [row.title, row.description, row.trainingAddress, row.location, row.company]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query),
    );
  }, [records, search]);

  const calendarEvents = useMemo<CalendarGridEvent[]>(
    () =>
      filtered.map((row) => ({
        id: row.id,
        title: row.title,
        start: row.eventDate,
        end: row.endDate,
        company: row.company,
        location: row.location,
      })),
    [filtered],
  );

  const selected = filtered.find((row) => row.id === selectedId) ?? null;

  function renderEventCard(row: CustomerEventRecord) {
    return (
      <article key={row.id} className={sectionStyles.card}>
        <h2>{row.title}</h2>
        <dl className={styles.eventMetaList}>
          <div>
            <dt>Date</dt>
            <dd>{row.eventDate ? formatDate(row.eventDate) : "Date TBC"}</dd>
          </div>
          <div>
            <dt>Time</dt>
            <dd>{formatTimeRange(row)}</dd>
          </div>
          <div>
            <dt>Duration</dt>
            <dd>{formatDuration(row.eventDate, row.endDate) ?? "Duration not set"}</dd>
          </div>
          <div>
            <dt>Training address</dt>
            <dd>{row.trainingAddress?.trim() || "—"}</dd>
          </div>
          <div>
            <dt>Location</dt>
            <dd>{row.location?.trim() || "—"}</dd>
          </div>
        </dl>
        <p className={sectionStyles.cardBody}>
          {row.description?.trim() || "No description provided."}
        </p>
      </article>
    );
  }

  return (
    <div>
      <CustomerPageHeader
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "Events" },
        ]}
        title="Events"
        subtitle="Your company’s customer-visible training events."
      />

      <p className={styles.companyMeta}>
        Calendar for <strong>{companyName}</strong>. Private events and other
        companies’ bookings are not included.
      </p>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search events, locations, descriptions…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
        <div className={styles.viewToggle} role="group" aria-label="View mode">
          {(["list", "month", "week"] as const).map((option) => (
            <button
              key={option}
              type="button"
              className={
                view === option ? styles.viewToggleActive : styles.viewToggleButton
              }
              aria-pressed={view === option}
              onClick={() => setView(option)}
            >
              {option[0].toUpperCase() + option.slice(1)}
            </button>
          ))}
        </div>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} event{records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No events found</h2>
          <p>
            {records.length === 0
              ? "No events have been shared with your company yet."
              : "Try adjusting your search to find matching events."}
          </p>
        </div>
      ) : view === "list" ? (
        <section className={styles.cardGrid} aria-label="Training events">
          {filtered.map(renderEventCard)}
        </section>
      ) : (
        <>
          <CalendarGrid
            events={calendarEvents}
            view={view}
            cursor={cursor}
            onCursorChange={setCursor}
            onEventClick={(event) => setSelectedId(event.id)}
          />
          <div className={styles.eventCalendarSelected}>
            {selected ? (
              renderEventCard(selected)
            ) : (
              <p className={styles.muted}>Select an event to view its details.</p>
            )}
          </div>
        </>
      )}
    </div>
  );
}

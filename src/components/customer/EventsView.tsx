"use client";

import { useMemo, useState } from "react";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type { CustomerEventRecord } from "@/types/models";

import styles from "./customer.module.css";
import sectionStyles from "./portalSections.module.css";

interface Props {
  companyName: string;
  records: CustomerEventRecord[];
}

function formatTime(value: string | null | undefined): string | null {
  if (!value?.trim()) {
    return null;
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return new Intl.DateTimeFormat("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatTimeRange(row: CustomerEventRecord): string {
  const start = formatTime(row.eventDate);
  const end = formatTime(row.endDate);
  if (start && end) return `${start} – ${end}`;
  if (start) return start;
  if (end) return end;
  return "Time TBC";
}

export function EventsView({ companyName, records }: Props) {
  const [search, setSearch] = useState("");
  const [view, setView] = useState<"list" | "calendar">("list");
  const [monthCursor, setMonthCursor] = useState(() => {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1);
  });
  const [selectedDay, setSelectedDay] = useState<string | null>(null);

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) {
      return records;
    }
    return records.filter((row) =>
      [
        row.title,
        row.description,
        row.trainingAddress,
        row.location,
        row.company,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query),
    );
  }, [records, search]);

  const calendarDays = useMemo(() => {
    const year = monthCursor.getFullYear();
    const month = monthCursor.getMonth();
    const firstDay = new Date(year, month, 1);
    const startPad = (firstDay.getDay() + 6) % 7; // Monday-first
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const cells: Array<{ key: string; date: Date | null; count: number }> = [];

    for (let i = 0; i < startPad; i += 1) {
      cells.push({ key: `pad-${i}`, date: null, count: 0 });
    }

    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = new Date(year, month, day);
      const key = date.toISOString().slice(0, 10);
      const count = filtered.filter((row) => {
        if (!row.eventDate) return false;
        const eventDay = new Date(row.eventDate);
        if (Number.isNaN(eventDay.getTime())) return false;
        return eventDay.toISOString().slice(0, 10) === key;
      }).length;
      cells.push({ key, date, count });
    }

    return cells;
  }, [filtered, monthCursor]);

  const selectedEvents = useMemo(() => {
    if (!selectedDay) return [];
    return filtered.filter((row) => {
      if (!row.eventDate) return false;
      const eventDay = new Date(row.eventDate);
      if (Number.isNaN(eventDay.getTime())) return false;
      return eventDay.toISOString().slice(0, 10) === selectedDay;
    });
  }, [filtered, selectedDay]);

  const monthLabel = new Intl.DateTimeFormat("en-GB", {
    month: "long",
    year: "numeric",
  }).format(monthCursor);

  function renderEventCard(row: CustomerEventRecord) {
    return (
      <article key={row.id} className={sectionStyles.card}>
        <h2>{row.title}</h2>
        <dl className={styles.eventMetaList}>
          <div>
            <dt>Date</dt>
            <dd>
              {row.eventDate ? formatDisplayDate(row.eventDate) : "Date TBC"}
            </dd>
          </div>
          <div>
            <dt>Time</dt>
            <dd>{formatTimeRange(row)}</dd>
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
        subtitle="Upcoming and scheduled training events for your company."
      />

      <p className={styles.companyMeta}>
        Showing events for <strong>{companyName}</strong>
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
          <button
            type="button"
            className={
              view === "list" ? styles.viewToggleActive : styles.viewToggleButton
            }
            onClick={() => setView("list")}
          >
            List
          </button>
          <button
            type="button"
            className={
              view === "calendar"
                ? styles.viewToggleActive
                : styles.viewToggleButton
            }
            onClick={() => setView("calendar")}
          >
            Calendar
          </button>
        </div>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} event
        {records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No events found</h2>
          <p>
            {records.length === 0
              ? "No upcoming events have been shared with your company yet."
              : "Try adjusting your search to find matching events."}
          </p>
        </div>
      ) : view === "list" ? (
        <section className={styles.cardGrid} aria-label="Training events">
          {filtered.map(renderEventCard)}
        </section>
      ) : (
        <div className={styles.eventCalendar}>
          <div className={styles.eventCalendarHeader}>
            <button
              type="button"
              className={styles.viewToggleButton}
              onClick={() =>
                setMonthCursor(
                  (current) =>
                    new Date(current.getFullYear(), current.getMonth() - 1, 1),
                )
              }
            >
              Previous
            </button>
            <h2>{monthLabel}</h2>
            <button
              type="button"
              className={styles.viewToggleButton}
              onClick={() =>
                setMonthCursor(
                  (current) =>
                    new Date(current.getFullYear(), current.getMonth() + 1, 1),
                )
              }
            >
              Next
            </button>
          </div>
          <div className={styles.eventCalendarWeekdays}>
            {["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"].map((day) => (
              <span key={day}>{day}</span>
            ))}
          </div>
          <div className={styles.eventCalendarGrid}>
            {calendarDays.map((cell) =>
              cell.date ? (
                <button
                  key={cell.key}
                  type="button"
                  className={
                    selectedDay === cell.key
                      ? styles.eventCalendarDayActive
                      : styles.eventCalendarDay
                  }
                  onClick={() => setSelectedDay(cell.key)}
                >
                  <span>{cell.date.getDate()}</span>
                  {cell.count > 0 ? (
                    <em>
                      {cell.count} event{cell.count === 1 ? "" : "s"}
                    </em>
                  ) : null}
                </button>
              ) : (
                <span key={cell.key} className={styles.eventCalendarPad} />
              ),
            )}
          </div>
          <div className={styles.eventCalendarSelected}>
            {selectedDay ? (
              selectedEvents.length > 0 ? (
                <section className={styles.cardGrid} aria-label="Selected day">
                  {selectedEvents.map(renderEventCard)}
                </section>
              ) : (
                <p className={styles.muted}>No events on this day.</p>
              )
            ) : (
              <p className={styles.muted}>Select a day to view events.</p>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

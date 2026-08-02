"use client";

import { useMemo } from "react";

import {
  formatDate,
  formatMonthYear,
  formatTime,
} from "@/lib/utils/formatDate";

import styles from "./calendarGrid.module.css";

export type CalendarView = "month" | "week";

export interface CalendarGridEvent {
  id: string;
  title: string;
  start: string | null;
  end?: string | null;
  company?: string | null;
  location?: string | null;
}

interface CalendarGridProps {
  events: CalendarGridEvent[];
  view: CalendarView;
  cursor: Date;
  onCursorChange: (date: Date) => void;
  onEventClick?: (event: CalendarGridEvent) => void;
  onSlotClick?: (date: Date) => void;
  emptyLabel?: string;
}

const WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const HOURS = Array.from({ length: 24 }, (_, hour) => hour);
const HOUR_HEIGHT = 48;

function startOfDay(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function startOfWeek(value: Date): Date {
  const date = startOfDay(value);
  date.setDate(date.getDate() - ((date.getDay() + 6) % 7));
  return date;
}

function addDays(value: Date, amount: number): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate() + amount);
}

function dateKey(value: Date): string {
  return [
    value.getFullYear(),
    String(value.getMonth() + 1).padStart(2, "0"),
    String(value.getDate()).padStart(2, "0"),
  ].join("-");
}

function parseEventDate(value: string | null | undefined): Date | null {
  if (!value?.trim()) return null;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function sameDay(left: Date, right: Date): boolean {
  return dateKey(left) === dateKey(right);
}

function eventTimeLabel(event: CalendarGridEvent): string {
  const start = formatTime(event.start);
  const end = formatTime(event.end);
  if (start && end) return `${start}–${end}`;
  return start ?? end ?? "Time TBC";
}

function EventChip({
  event,
  onClick,
  compact = false,
}: {
  event: CalendarGridEvent;
  onClick?: (event: CalendarGridEvent) => void;
  compact?: boolean;
}) {
  const label = `${event.title}, ${eventTimeLabel(event)}${
    event.company ? `, ${event.company}` : ""
  }`;
  return (
    <button
      type="button"
      className={compact ? styles.weekEvent : styles.monthEvent}
      title={label}
      aria-label={label}
      onClick={(clickEvent) => {
        clickEvent.stopPropagation();
        onClick?.(event);
      }}
    >
      <strong>{event.title}</strong>
      <span>{eventTimeLabel(event)}</span>
    </button>
  );
}

export function CalendarGrid({
  events,
  view,
  cursor,
  onCursorChange,
  onEventClick,
  onSlotClick,
  emptyLabel = "No events in this period.",
}: CalendarGridProps) {
  const today = useMemo(() => startOfDay(new Date()), []);
  const parsedEvents = useMemo(
    () =>
      events
        .map((event) => ({ event, start: parseEventDate(event.start) }))
        .filter(
          (entry): entry is { event: CalendarGridEvent; start: Date } =>
            entry.start !== null,
        )
        .sort((a, b) => a.start.getTime() - b.start.getTime()),
    [events],
  );

  const weekStart = startOfWeek(cursor);
  const weekDays = Array.from({ length: 7 }, (_, index) =>
    addDays(weekStart, index),
  );

  const monthDays = useMemo(() => {
    const first = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
    const gridStart = addDays(first, -((first.getDay() + 6) % 7));
    return Array.from({ length: 42 }, (_, index) => addDays(gridStart, index));
  }, [cursor]);

  function move(amount: number) {
    onCursorChange(
      view === "month"
        ? new Date(cursor.getFullYear(), cursor.getMonth() + amount, 1)
        : addDays(cursor, amount * 7),
    );
  }

  const periodLabel =
    view === "month"
      ? formatMonthYear(cursor)
      : `${formatDate(weekDays[0])} – ${formatDate(weekDays[6])}`;

  return (
    <section className={styles.calendar} aria-label={`${periodLabel} calendar`}>
      <div className={styles.toolbar}>
        <div className={styles.navigation}>
          <button type="button" onClick={() => move(-1)} aria-label="Previous period">
            ‹
          </button>
          <button type="button" onClick={() => onCursorChange(new Date())}>
            Today
          </button>
          <button type="button" onClick={() => move(1)} aria-label="Next period">
            ›
          </button>
        </div>
        <h2>{periodLabel}</h2>
        <span className={styles.eventCount}>
          {events.length} event{events.length === 1 ? "" : "s"}
        </span>
      </div>

      {view === "month" ? (
        <div className={styles.monthScroller}>
          <div className={styles.monthGrid}>
            {WEEKDAYS.map((day) => (
              <div key={day} className={styles.weekday}>
                {day}
              </div>
            ))}
            {monthDays.map((day) => {
              const dayEvents = parsedEvents
                .filter((entry) => sameDay(entry.start, day))
                .map((entry) => entry.event);
              const isCurrentMonth = day.getMonth() === cursor.getMonth();
              const classNames = [styles.monthDay];
              if (!isCurrentMonth) classNames.push(styles.outsideMonth);
              if (sameDay(day, today)) classNames.push(styles.today);
              return (
                <div
                  key={dateKey(day)}
                  className={classNames.join(" ")}
                  role={onSlotClick ? "button" : undefined}
                  tabIndex={onSlotClick ? 0 : undefined}
                  aria-label={`${formatDate(day)}${dayEvents.length ? `, ${dayEvents.length} events` : ""}`}
                  onClick={() => onSlotClick?.(new Date(day.getFullYear(), day.getMonth(), day.getDate(), 9))}
                  onKeyDown={(event) => {
                    if (onSlotClick && (event.key === "Enter" || event.key === " ")) {
                      event.preventDefault();
                      onSlotClick(new Date(day.getFullYear(), day.getMonth(), day.getDate(), 9));
                    }
                  }}
                >
                  <span className={styles.dayNumber}>{day.getDate()}</span>
                  <div className={styles.monthEvents}>
                    {dayEvents.slice(0, 3).map((event) => (
                      <EventChip key={event.id} event={event} onClick={onEventClick} />
                    ))}
                    {dayEvents.length > 3 ? (
                      <span className={styles.moreCount}>+{dayEvents.length - 3} more</span>
                    ) : null}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ) : (
        <div className={styles.weekScroller}>
          <div className={styles.weekGrid}>
            <div className={styles.timeCorner} />
            {weekDays.map((day) => (
              <div
                key={dateKey(day)}
                className={`${styles.weekDayHeader} ${sameDay(day, today) ? styles.weekToday : ""}`}
              >
                <span>{WEEKDAYS[(day.getDay() + 6) % 7]}</span>
                <strong>{formatDate(day)}</strong>
              </div>
            ))}
            <div className={styles.timeAxis}>
              {HOURS.map((hour) => (
                <span key={hour}>{String(hour).padStart(2, "0")}:00</span>
              ))}
            </div>
            {weekDays.map((day) => {
              const dayEvents = parsedEvents.filter((entry) => sameDay(entry.start, day));
              return (
                <div key={dateKey(day)} className={styles.weekDayColumn}>
                  {HOURS.map((hour) => (
                    <button
                      key={hour}
                      type="button"
                      className={styles.hourSlot}
                      aria-label={`Create event ${formatDate(day)} at ${String(hour).padStart(2, "0")}:00`}
                      onClick={() => onSlotClick?.(new Date(day.getFullYear(), day.getMonth(), day.getDate(), hour))}
                      disabled={!onSlotClick}
                    />
                  ))}
                  {dayEvents.map(({ event, start }) => {
                    const end = parseEventDate(event.end);
                    const durationMinutes = end
                      ? Math.max(30, (end.getTime() - start.getTime()) / 60_000)
                      : 60;
                    return (
                      <div
                        key={event.id}
                        className={styles.weekEventPosition}
                        style={{
                          top: `${(start.getHours() + start.getMinutes() / 60) * HOUR_HEIGHT}px`,
                          height: `${Math.max(24, (durationMinutes / 60) * HOUR_HEIGHT)}px`,
                        }}
                      >
                        <EventChip event={event} onClick={onEventClick} compact />
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {parsedEvents.length === 0 ? <p className={styles.empty}>{emptyLabel}</p> : null}
    </section>
  );
}

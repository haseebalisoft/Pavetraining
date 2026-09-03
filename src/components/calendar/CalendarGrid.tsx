"use client";

import { useEffect, useMemo, useState } from "react";

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
const WEEKDAY_SHORT = ["M", "T", "W", "T", "F", "S", "S"];
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
  const dateOnly = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim());
  if (dateOnly) {
    return new Date(
      Number(dateOnly[1]),
      Number(dateOnly[2]) - 1,
      Number(dateOnly[3]),
    );
  }
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

function eventTouchesDay(
  start: Date,
  end: Date | null,
  day: Date,
): boolean {
  const dayStart = startOfDay(day);
  const dayEnd = addDays(dayStart, 1);
  const eventEnd = end && end.getTime() >= start.getTime() ? end : start;
  return start < dayEnd && eventEnd >= dayStart;
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
  const [selectedDay, setSelectedDay] = useState(() => startOfDay(new Date()));

  const parsedEvents = useMemo(
    () =>
      events
        .map((event) => ({
          event,
          start: parseEventDate(event.start),
          end: parseEventDate(event.end ?? null),
        }))
        .filter(
          (
            entry,
          ): entry is {
            event: CalendarGridEvent;
            start: Date;
            end: Date | null;
          } => entry.start !== null,
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

  const cursorMonthKey = `${cursor.getFullYear()}-${cursor.getMonth()}`;

  useEffect(() => {
    if (view !== "month") return;

    const todayInMonth =
      today.getMonth() === cursor.getMonth() &&
      today.getFullYear() === cursor.getFullYear();
    if (todayInMonth) {
      setSelectedDay(today);
      return;
    }

    const firstWithEvents = monthDays.find(
      (day) =>
        day.getMonth() === cursor.getMonth() &&
        parsedEvents.some((entry) =>
          eventTouchesDay(entry.start, entry.end, day),
        ),
    );
    setSelectedDay(
      firstWithEvents ??
        new Date(cursor.getFullYear(), cursor.getMonth(), 1),
    );
    // Reset selection when the visible month changes, not on every day tap.
    // eslint-disable-next-line react-hooks/exhaustive-deps -- intentional month-scoped reset
  }, [view, cursorMonthKey]);

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

  const selectedDayEvents = useMemo(
    () =>
      parsedEvents
        .filter((entry) =>
          eventTouchesDay(entry.start, entry.end, selectedDay),
        )
        .map((entry) => entry.event),
    [parsedEvents, selectedDay],
  );

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
        <>
          <div className={styles.monthScroller}>
            <div className={styles.monthGrid}>
              {WEEKDAYS.map((day, index) => (
                <div key={day} className={styles.weekday} aria-label={day}>
                  <span className={styles.weekdayFull} aria-hidden="true">
                    {day}
                  </span>
                  <span className={styles.weekdayShort} aria-hidden="true">
                    {WEEKDAY_SHORT[index]}
                  </span>
                </div>
              ))}
              {monthDays.map((day) => {
                const dayEvents = parsedEvents
                  .filter((entry) =>
                    eventTouchesDay(entry.start, entry.end, day),
                  )
                  .map((entry) => entry.event);
                const isCurrentMonth = day.getMonth() === cursor.getMonth();
                const isSelected = sameDay(day, selectedDay);
                const classNames = [styles.monthDay];
                if (!isCurrentMonth) classNames.push(styles.outsideMonth);
                if (sameDay(day, today)) classNames.push(styles.today);
                if (isSelected) classNames.push(styles.selectedDay);
                return (
                  <div
                    key={dateKey(day)}
                    className={classNames.join(" ")}
                    role="button"
                    tabIndex={0}
                    aria-pressed={isSelected}
                    aria-label={`${formatDate(day)}${dayEvents.length ? `, ${dayEvents.length} events` : ""}`}
                    onClick={() => {
                      setSelectedDay(startOfDay(day));
                      if (onSlotClick && dayEvents.length === 0) {
                        onSlotClick(
                          new Date(
                            day.getFullYear(),
                            day.getMonth(),
                            day.getDate(),
                            9,
                          ),
                        );
                      }
                    }}
                    onKeyDown={(event) => {
                      if (event.key !== "Enter" && event.key !== " ") return;
                      event.preventDefault();
                      setSelectedDay(startOfDay(day));
                      if (onSlotClick && dayEvents.length === 0) {
                        onSlotClick(
                          new Date(
                            day.getFullYear(),
                            day.getMonth(),
                            day.getDate(),
                            9,
                          ),
                        );
                      }
                    }}
                  >
                    <span className={styles.dayNumber}>{day.getDate()}</span>
                    <div className={styles.monthEvents}>
                      {dayEvents.slice(0, 3).map((event) => (
                        <EventChip
                          key={`${event.id}-${dateKey(day)}`}
                          event={event}
                          onClick={onEventClick}
                        />
                      ))}
                      {dayEvents.length > 3 ? (
                        <button
                          type="button"
                          className={styles.moreCount}
                          onClick={(clickEvent) => {
                            clickEvent.stopPropagation();
                            setSelectedDay(startOfDay(day));
                            onEventClick?.(dayEvents[3] ?? dayEvents[0]);
                          }}
                        >
                          +{dayEvents.length - 3} more
                        </button>
                      ) : null}
                    </div>
                    {dayEvents.length > 0 ? (
                      <div
                        className={styles.eventDots}
                        aria-hidden="true"
                      >
                        {dayEvents.slice(0, 3).map((event) => (
                          <span
                            key={`${event.id}-dot-${dateKey(day)}`}
                            className={styles.eventDot}
                          />
                        ))}
                        {dayEvents.length > 3 ? (
                          <span className={styles.eventDotMore}>
                            +{dayEvents.length - 3}
                          </span>
                        ) : null}
                      </div>
                    ) : null}
                  </div>
                );
              })}
            </div>
          </div>

          <div className={styles.dayAgenda} aria-live="polite">
            <div className={styles.dayAgendaHeader}>
              <h3>{formatDate(selectedDay)}</h3>
              <span>
                {selectedDayEvents.length} event
                {selectedDayEvents.length === 1 ? "" : "s"}
              </span>
            </div>
            {selectedDayEvents.length === 0 ? (
              <p className={styles.dayAgendaEmpty}>
                No events on this day. Tap another date, or switch to List for
                all events.
              </p>
            ) : (
              <ul className={styles.dayAgendaList}>
                {selectedDayEvents.map((event) => (
                  <li key={event.id}>
                    <button
                      type="button"
                      className={styles.dayAgendaItem}
                      onClick={() => onEventClick?.(event)}
                    >
                      <strong>{event.title}</strong>
                      <span>{eventTimeLabel(event)}</span>
                      {event.location ? (
                        <span className={styles.dayAgendaMeta}>
                          {event.location}
                        </span>
                      ) : event.company ? (
                        <span className={styles.dayAgendaMeta}>
                          {event.company}
                        </span>
                      ) : null}
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </>
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
              const dayEvents = parsedEvents.filter((entry) =>
                eventTouchesDay(entry.start, entry.end, day),
              );
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
                  {dayEvents.map(({ event, start, end }) => {
                    const isStartDay = sameDay(start, day);
                    const durationMinutes = end
                      ? Math.max(30, (end.getTime() - start.getTime()) / 60_000)
                      : 60;
                    const top = isStartDay
                      ? (start.getHours() + start.getMinutes() / 60) * HOUR_HEIGHT
                      : 0;
                    const height = isStartDay
                      ? Math.max(24, (durationMinutes / 60) * HOUR_HEIGHT)
                      : HOUR_HEIGHT * 24;
                    return (
                      <div
                        key={`${event.id}-${dateKey(day)}`}
                        className={styles.weekEventPosition}
                        style={{
                          top: `${top}px`,
                          height: `${Math.min(HOUR_HEIGHT * 24, height)}px`,
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

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

function place(row: CustomerEventRecord): string | null {
  return row.trainingAddress?.trim() || row.location?.trim() || null;
}

export function EventsView({ companyName, records }: Props) {
  const [search, setSearch] = useState("");

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) {
      return records;
    }
    return records.filter((row) =>
      [row.title, row.description, row.company, place(row)]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query),
    );
  }, [records, search]);

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
              ? "There are no customer-visible events for your company yet."
              : "Try adjusting your search to find matching events."}
          </p>
        </div>
      ) : (
        <section className={styles.cardGrid} aria-label="Training events">
          {filtered.map((row) => {
            const time = formatTime(row.eventDate);
            const endTime = formatTime(row.endDate);
            const location = place(row);

            return (
              <article key={row.id} className={sectionStyles.card}>
                <h2>{row.title}</h2>
                <p className={sectionStyles.cardMeta}>
                  {row.eventDate ? formatDisplayDate(row.eventDate) : "Date TBC"}
                  {time ? ` · ${time}` : ""}
                  {endTime ? ` – ${endTime}` : ""}
                </p>
                {location ? (
                  <p className={sectionStyles.cardMeta}>{location}</p>
                ) : null}
                <p className={sectionStyles.cardBody}>
                  {row.description?.trim() || "No description provided."}
                </p>
              </article>
            );
          })}
        </section>
      )}
    </div>
  );
}

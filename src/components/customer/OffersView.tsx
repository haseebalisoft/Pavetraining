"use client";

import { useMemo, useState } from "react";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import { StatusBadge } from "@/components/ui/StatusBadge";
import { formatDisplayDate } from "@/lib/training/expiryFilters";
import type { CustomerOfferRecord } from "@/types/models";

import styles from "./customer.module.css";
import sectionStyles from "./portalSections.module.css";

interface Props {
  companyName: string;
  records: CustomerOfferRecord[];
}

export function OffersView({ companyName, records }: Props) {
  const [search, setSearch] = useState("");

  const filtered = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) {
      return records;
    }
    return records.filter((row) =>
      [row.title, row.description, row.status]
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
          { label: "Offers" },
        ]}
        title="Offers"
        subtitle="Active promotions and offers available to your organisation."
      />

      <p className={styles.companyMeta}>
        Showing offers for <strong>{companyName}</strong>
      </p>

      <div className={styles.toolbar}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>Search</span>
          <input
            className={styles.input}
            type="search"
            placeholder="Search offers…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </label>
      </div>

      <p className={styles.resultCount}>
        {filtered.length} of {records.length} offer
        {records.length === 1 ? "" : "s"}
      </p>

      {filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <h2>No offers found</h2>
          <p>
            {records.length === 0
              ? "There are no active customer-visible offers for your company yet."
              : "Try adjusting your search to find matching offers."}
          </p>
        </div>
      ) : (
        <section className={styles.cardGrid} aria-label="Offers and promotions">
          {filtered.map((row) => (
            <article key={row.id} className={sectionStyles.card}>
              <StatusBadge
                label={row.status ?? "Active"}
                tone={
                  (row.status ?? "Active").toLowerCase() === "active"
                    ? "ok"
                    : "neutral"
                }
              />
              <h2 style={{ marginTop: "0.55rem" }}>{row.title}</h2>
              <p className={sectionStyles.cardMeta}>
                {row.startDate ? formatDisplayDate(row.startDate) : "Start TBC"}
                {" – "}
                {row.endDate ? formatDisplayDate(row.endDate) : "End TBC"}
              </p>
              <p className={sectionStyles.cardBody}>
                {row.description?.trim() || "No description provided."}
              </p>
            </article>
          ))}
        </section>
      )}
    </div>
  );
}

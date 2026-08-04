"use client";

import { useMemo, useState } from "react";
import Link from "next/link";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import type { CustomerDocumentRecord } from "@/types/models";
import { formatDate } from "@/lib/utils/formatDate";
import styles from "./customer.module.css";

export type CustomerCourseItem = {
  code: string;
  title: string;
  group?: string | null;
};

type TabId = "courses" | "npors-docs";

interface Props {
  companyName: string;
  courses: CustomerCourseItem[];
  resourceDocs: CustomerDocumentRecord[];
}

function isResourceDoc(doc: CustomerDocumentRecord): boolean {
  const blob = `${doc.name} ${doc.documentType ?? ""}`.toLowerCase();
  return (
    blob.includes("objective") ||
    blob.includes("site requirement") ||
    blob.includes("site requirements") ||
    blob.includes("course objective") ||
    (blob.includes("npors") &&
      (blob.includes("objective") ||
        blob.includes("requirement") ||
        blob.includes("brochure")))
  );
}

export function CustomerCoursesView({
  companyName,
  courses,
  resourceDocs,
}: Props) {
  const [tab, setTab] = useState<TabId>("courses");
  const [search, setSearch] = useState("");

  const filteredCourses = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return courses;
    return courses.filter((row) =>
      `${row.code} ${row.title} ${row.group ?? ""}`.toLowerCase().includes(q),
    );
  }, [courses, search]);

  const docs = useMemo(() => {
    const matched = resourceDocs.filter(isResourceDoc);
    return matched.length > 0 ? matched : resourceDocs;
  }, [resourceDocs]);

  return (
    <div>
      <CustomerPageHeader
        title="Training Delivery"
        subtitle="Browse PAVE brochures and course information for NPORS, EUSR, Streetworks, and NVQ — plus objectives and site requirements documents."
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "Training Delivery" },
        ]}
      />

      <p className={styles.companyMeta}>
        Shared with <strong>{companyName}</strong>
      </p>

      <div className={styles.viewToggle} role="tablist" aria-label="Training Delivery sections">
        <button
          type="button"
          role="tab"
          aria-selected={tab === "courses"}
          className={
            tab === "courses" ? styles.viewToggleActive : styles.viewToggleButton
          }
          onClick={() => setTab("courses")}
        >
          Course catalogue
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={tab === "npors-docs"}
          className={
            tab === "npors-docs"
              ? styles.viewToggleActive
              : styles.viewToggleButton
          }
          onClick={() => setTab("npors-docs")}
        >
          NPORS objectives &amp; site requirements
        </button>
      </div>

      {tab === "courses" ? (
        <>
          <div className={styles.toolbar} style={{ marginTop: "1rem" }}>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Search courses</span>
              <input
                className={styles.input}
                type="search"
                placeholder="Search by code or title (e.g. N001, Telehandler)…"
                value={search}
                onChange={(event) => setSearch(event.target.value)}
              />
            </label>
          </div>
          <p className={styles.resultCount}>
            {filteredCourses.length} of {courses.length} course
            {courses.length === 1 ? "" : "s"}
          </p>
          {filteredCourses.length === 0 ? (
            <div className={styles.emptyState}>
              <h2>No courses found</h2>
              <p>Try a different search, or contact Support to book training.</p>
            </div>
          ) : (
            <div className={styles.tableWrap}>
              <table className={styles.dataTable}>
                <thead>
                  <tr>
                    <th scope="col">Code</th>
                    <th scope="col">Course</th>
                    <th scope="col">Group</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredCourses.map((row) => (
                    <tr key={row.code}>
                      <td>{row.code}</td>
                      <td>{row.title}</td>
                      <td>{row.group?.trim() || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <p className={styles.companyMeta} style={{ marginTop: "1.25rem" }}>
            Ready to book?{" "}
            <Link className={styles.link} href="/customer/support">
              Contact PAVE Support
            </Link>{" "}
            or check{" "}
            <Link className={styles.link} href="/customer/events">
              Events / Bookings
            </Link>
            .
          </p>
        </>
      ) : (
        <section style={{ marginTop: "1.25rem" }} aria-label="NPORS documents">
          <div className={styles.supportCard} style={{ maxWidth: "none" }}>
            <h2 className={styles.supportTitle}>
              NPORS course objectives &amp; site requirements
            </h2>
            <p className={styles.supportCopy}>
              Download the objectives and site requirement packs your team needs
              before attending. Documents appear here when PAVE uploads them to
              your company Documents library (names/types containing NPORS,
              objectives, or site requirements).
            </p>
          </div>

          {docs.length === 0 ? (
            <div className={styles.emptyState} style={{ marginTop: "1rem" }}>
              <h2>No documents uploaded yet</h2>
              <p>
                Ask PAVE to upload NPORS course objectives and site requirements
                into Documents, or{" "}
                <Link className={styles.link} href="/customer/support">
                  contact Support
                </Link>
                .
              </p>
            </div>
          ) : (
            <div className={styles.tableWrap} style={{ marginTop: "1rem" }}>
              <table className={styles.dataTable}>
                <thead>
                  <tr>
                    <th scope="col">Document</th>
                    <th scope="col">Type</th>
                    <th scope="col">Updated</th>
                    <th scope="col">View</th>
                    <th scope="col">Download</th>
                  </tr>
                </thead>
                <tbody>
                  {docs.map((row) => (
                    <tr key={row.id}>
                      <td>{row.name}</td>
                      <td>{row.documentType?.trim() || "—"}</td>
                      <td>
                        {row.uploadedDate ? formatDate(row.uploadedDate) : "—"}
                      </td>
                      <td>
                        {row.viewPath ? (
                          <a
                            className={styles.link}
                            href={row.viewPath}
                            target="_blank"
                            rel="noreferrer"
                          >
                            View
                          </a>
                        ) : (
                          "—"
                        )}
                      </td>
                      <td>
                        {row.canDownload && row.downloadPath ? (
                          <a className={styles.link} href={row.downloadPath}>
                            Download
                          </a>
                        ) : (
                          "—"
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <p className={styles.companyMeta} style={{ marginTop: "1rem" }}>
            All company files also live under{" "}
            <Link className={styles.link} href="/customer/documents">
              Documents
            </Link>
            .
          </p>
        </section>
      )}
    </div>
  );
}

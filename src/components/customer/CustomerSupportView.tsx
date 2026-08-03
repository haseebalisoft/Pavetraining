"use client";

import { useState, type FormEvent } from "react";

import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import styles from "./customer.module.css";

const CONTACT = {
  email: "info@pavetraining.co.uk",
  phone: "01606 351240",
  web: "https://www.pavetraining.co.uk",
};

interface Props {
  companyName: string;
  defaultEmail: string;
  defaultName?: string | null;
}

export function CustomerSupportView({
  companyName,
  defaultEmail,
  defaultName = "",
}: Props) {
  const [name, setName] = useState(defaultName ?? "");
  const [email, setEmail] = useState(defaultEmail);
  const [phone, setPhone] = useState("");
  const [subject, setSubject] = useState("");
  const [message, setMessage] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [sent, setSent] = useState(false);

  async function onSubmit(event: FormEvent) {
    event.preventDefault();
    setSaving(true);
    setError(null);
    setSent(false);
    try {
      const response = await fetch("/api/customer/support", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name,
          email,
          phone,
          companyName,
          subject,
          message,
        }),
      });
      if (!response.ok) {
        throw new Error(await readPublicApiError(response));
      }
      setSent(true);
      setSubject("");
      setMessage("");
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Something went wrong sending your message.",
      );
    } finally {
      setSaving(false);
    }
  }

  return (
    <div>
      <CustomerPageHeader
        title="Support"
        subtitle="Contact PAVE Training — send a message or use the details below."
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "Support" },
        ]}
      />

      <div className={styles.supportLayout}>
        <section className={styles.supportCard} aria-label="Contact details">
          <h2 className={styles.supportTitle}>Contact details</h2>
          <p className={styles.supportCopy}>
            Same team that looks after your training bookings and cards.
          </p>
          <ul className={styles.supportContactList}>
            <li>
              <span className={styles.supportContactLabel}>Email</span>
              <a className={styles.link} href={`mailto:${CONTACT.email}`}>
                {CONTACT.email}
              </a>
            </li>
            <li>
              <span className={styles.supportContactLabel}>Phone</span>
              <a className={styles.link} href={`tel:${CONTACT.phone.replace(/\s/g, "")}`}>
                {CONTACT.phone}
              </a>
            </li>
            <li>
              <span className={styles.supportContactLabel}>Website</span>
              <a
                className={styles.link}
                href={CONTACT.web}
                target="_blank"
                rel="noreferrer"
              >
                pavetraining.co.uk
              </a>
            </li>
            <li>
              <span className={styles.supportContactLabel}>Your company</span>
              <span>{companyName}</span>
            </li>
          </ul>
        </section>

        <section className={styles.supportCard} aria-label="Contact form">
          <h2 className={styles.supportTitle}>Send a message</h2>
          <p className={styles.supportCopy}>
            Tell us about renewals, bookings, certificates, or portal access.
            We will reply to the email you enter below.
          </p>

          {sent ? (
            <p className={styles.supportSuccess} role="status">
              Thanks — your message has been sent to PAVE Training.
            </p>
          ) : null}
          {error ? (
            <p className={styles.supportError} role="alert">
              {error}
            </p>
          ) : null}

          <form className={styles.supportForm} onSubmit={onSubmit}>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Your name</span>
              <input
                className={styles.input}
                value={name}
                onChange={(e) => setName(e.target.value)}
                required
                autoComplete="name"
              />
            </label>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Email</span>
              <input
                className={styles.input}
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
                autoComplete="email"
              />
            </label>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Phone (optional)</span>
              <input
                className={styles.input}
                type="tel"
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                autoComplete="tel"
              />
            </label>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Subject</span>
              <input
                className={styles.input}
                value={subject}
                onChange={(e) => setSubject(e.target.value)}
                required
                placeholder="e.g. Book NPORS refresher"
              />
            </label>
            <label className={styles.field}>
              <span className={styles.fieldLabel}>Message</span>
              <textarea
                className={styles.textarea}
                value={message}
                onChange={(e) => setMessage(e.target.value)}
                required
                rows={6}
                placeholder="Include candidate names, course codes, or dates if you have them."
              />
            </label>
            <button
              type="submit"
              className={styles.primaryButton}
              disabled={saving}
            >
              {saving ? "Sending…" : "Send message"}
            </button>
          </form>
        </section>
      </div>
    </div>
  );
}

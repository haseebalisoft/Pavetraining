"use client";

import { useCallback, useEffect, useState } from "react";

import { useAdminToast } from "@/components/admin/AdminToast";
import { Breadcrumbs } from "@/components/ui/Breadcrumbs";
import { LoadingState } from "@/components/ui/States";
import { readPublicApiError } from "@/lib/errors/publicMessages";
import { formatDateTime } from "@/lib/utils/formatDate";
import type {
  ExpiryNotificationRunResult,
  NotificationLogEntry,
  NotificationSettingsSummary,
  NotificationSendResult,
} from "@/types/notifications";

import styles from "@/components/admin/admin.module.css";

type LogsPayload = {
  settings: NotificationSettingsSummary;
  logs: NotificationLogEntry[];
  failedLogs: NotificationLogEntry[];
  documentStatus: Array<{
    id: string;
    name: string;
    company: string | null;
    documentType: string | null;
    customerVisible: boolean;
    notifyCustomer: boolean;
    notificationSent: boolean;
  }>;
};

export function AdminNotificationsClient() {
  const { pushToast } = useAdminToast();
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [data, setData] = useState<LogsPayload | null>(null);
  const [testTo, setTestTo] = useState("");
  const [expiryResult, setExpiryResult] =
    useState<ExpiryNotificationRunResult | null>(null);
  const [lastTest, setLastTest] = useState<NotificationSendResult | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const response = await fetch("/api/admin/notifications/logs", {
        cache: "no-store",
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const payload = (await response.json()) as LogsPayload;
      setData(payload);
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Failed to load notifications",
        "error",
      );
    } finally {
      setLoading(false);
    }
  }, [pushToast]);

  useEffect(() => {
    void load();
  }, [load]);

  async function sendTest() {
    setBusy(true);
    try {
      const response = await fetch("/api/admin/notifications/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ to: testTo.trim() || undefined }),
      });
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const payload = (await response.json()) as {
        result: NotificationSendResult;
      };
      setLastTest(payload.result);
      pushToast(
        payload.result.status === "sent"
          ? "Test email sent."
          : payload.result.errorMessage ||
              `Test status: ${payload.result.status}`,
        payload.result.status === "failed" ? "error" : "success",
      );
      await load();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Test email failed",
        "error",
      );
    } finally {
      setBusy(false);
    }
  }

  async function runExpiry(dryRun: boolean) {
    setBusy(true);
    try {
      const response = await fetch(
        "/api/admin/notifications/run-expiry-check",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ dryRun }),
        },
      );
      if (!response.ok) throw new Error(await readPublicApiError(response));
      const payload = (await response.json()) as {
        result: ExpiryNotificationRunResult;
      };
      setExpiryResult(payload.result);
      pushToast(
        dryRun
          ? `Dry run prepared ${payload.result.remindersPrepared} reminder(s).`
          : `Expiry check: sent ${payload.result.sent}, failed ${payload.result.failed}.`,
        payload.result.failed > 0 ? "error" : "success",
      );
      await load();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Expiry check failed",
        "error",
      );
    } finally {
      setBusy(false);
    }
  }

  async function resendDocument(documentId: string) {
    setBusy(true);
    try {
      const response = await fetch(
        `/api/admin/notifications/send-document/${documentId}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ force: true }),
        },
      );
      if (!response.ok) throw new Error(await readPublicApiError(response));
      pushToast("Document notification triggered.", "success");
      await load();
    } catch (error) {
      pushToast(
        error instanceof Error ? error.message : "Document notify failed",
        "error",
      );
    } finally {
      setBusy(false);
    }
  }

  if (loading && !data) {
    return <LoadingState label="Loading notifications…" />;
  }

  const settings = data?.settings;

  return (
    <div>
      <header className={styles.pageHeader}>
        <div>
          <Breadcrumbs
            items={[
              { label: "Admin", href: "/admin" },
              { label: "Notifications" },
            ]}
          />
          <p className={styles.eyebrow}>Admin</p>
          <h1 className={styles.title}>Notifications</h1>
          <p className={styles.subtitle}>
            Document upload notices, expiry reminders, and delivery logs. Customer
            preferences stay on the Permissions List (admin-managed).
          </p>
        </div>
      </header>

      {settings ? (
        <section className={styles.syncPanel}>
          <div className={styles.syncPanelHeader}>
            <strong>Settings summary</strong>
            <button
              type="button"
              className={styles.linkButton}
              onClick={() => void load()}
              disabled={busy}
            >
              Refresh
            </button>
          </div>
          <div className={styles.settingsGrid}>
            <div>
              <span className={styles.fieldLabel}>Notifications</span>
              <p>{settings.notificationsEnabled ? "Enabled" : "Disabled"}</p>
            </div>
            <div>
              <span className={styles.fieldLabel}>Expiry reminders</span>
              <p>
                {settings.expiryRemindersEnabled ? "Enabled" : "Disabled"}
                {settings.expiredRemindersEnabled ? " · expired on" : ""}
              </p>
            </div>
            <div>
              <span className={styles.fieldLabel}>6-month window</span>
              <p>91–{settings.expiry6mMaxDays} days</p>
            </div>
            <div>
              <span className={styles.fieldLabel}>From email</span>
              <p>{settings.fromEmail ?? "Not configured"}</p>
            </div>
            <div>
              <span className={styles.fieldLabel}>Email delivery</span>
              <p>
                {settings.emailConfigured
                  ? "Graph sendMail ready"
                  : "Not sent - email not configured"}
              </p>
            </div>
            <div>
              <span className={styles.fieldLabel}>Portal URL</span>
              <p>{settings.portalUrl ?? "—"}</p>
            </div>
          </div>
        </section>
      ) : null}

      <section className={styles.syncPanel}>
        <div className={styles.syncPanelHeader}>
          <strong>Actions</strong>
        </div>
        <div className={styles.syncToolbar}>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>Test email to</span>
            <input
              className={styles.input}
              type="email"
              placeholder="defaults to your login email"
              value={testTo}
              onChange={(event) => setTestTo(event.target.value)}
              disabled={busy}
            />
          </label>
          <button
            type="button"
            className={styles.primaryButton}
            disabled={busy}
            onClick={() => void sendTest()}
          >
            Send test email
          </button>
          <button
            type="button"
            className={styles.secondaryButton}
            disabled={busy}
            onClick={() => void runExpiry(true)}
          >
            Dry-run expiry check
          </button>
          <button
            type="button"
            className={styles.secondaryButton}
            disabled={busy}
            onClick={() => void runExpiry(false)}
          >
            Run expiry check
          </button>
        </div>
        {lastTest ? (
          <p className={styles.mutedNote}>
            Last test: {lastTest.status}
            {lastTest.errorMessage ? ` — ${lastTest.errorMessage}` : ""}
          </p>
        ) : null}
      </section>

      {expiryResult ? (
        <section className={styles.syncPanel}>
          <div className={styles.syncPanelHeader}>
            <strong>
              Expiry check {expiryResult.dryRun ? "(dry run)" : "results"}
            </strong>
            <span>
              Prepared {expiryResult.remindersPrepared} · Sent{" "}
              {expiryResult.sent} · Failed {expiryResult.failed} · Skipped{" "}
              {expiryResult.skipped} · Not configured{" "}
              {expiryResult.notConfigured}
            </span>
          </div>
          <div className={styles.tableWrap}>
            <table className={styles.dataTable}>
              <thead>
                <tr>
                  <th scope="col">Company</th>
                  <th scope="col">Recipient</th>
                  <th scope="col">Window</th>
                  <th scope="col">Candidates</th>
                  <th scope="col">Status</th>
                  <th scope="col">Detail</th>
                </tr>
              </thead>
              <tbody>
                {expiryResult.items.length === 0 ? (
                  <tr>
                    <td colSpan={6}>No expiry reminders matched.</td>
                  </tr>
                ) : (
                  expiryResult.items.map((item, index) => (
                    <tr
                      key={`${item.recipientEmail}-${item.window}-${index}`}
                    >
                      <td>{item.companyName}</td>
                      <td>{item.recipientEmail}</td>
                      <td>{item.window}</td>
                      <td>{item.candidateCount}</td>
                      <td>{item.status}</td>
                      <td>{item.errorMessage ?? "—"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>
      ) : null}

      <section className={styles.syncPanel}>
        <div className={styles.syncPanelHeader}>
          <strong>Failed notifications</strong>
          <span>{data?.failedLogs.length ?? 0}</span>
        </div>
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">When</th>
                <th scope="col">Type</th>
                <th scope="col">Recipient</th>
                <th scope="col">Subject</th>
                <th scope="col">Error</th>
              </tr>
            </thead>
            <tbody>
              {(data?.failedLogs.length ?? 0) === 0 ? (
                <tr>
                  <td colSpan={5}>No failed notifications logged.</td>
                </tr>
              ) : (
                data!.failedLogs.map((entry) => (
                  <tr key={entry.id}>
                    <td>{formatDateTime(entry.createdAt)}</td>
                    <td>{entry.type}</td>
                    <td>{entry.recipientEmail ?? "—"}</td>
                    <td>{entry.subject}</td>
                    <td>{entry.errorMessage ?? "—"}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>

      <section className={styles.syncPanel}>
        <div className={styles.syncPanelHeader}>
          <strong>Recent notification log</strong>
        </div>
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">When</th>
                <th scope="col">Type</th>
                <th scope="col">Status</th>
                <th scope="col">Recipient</th>
                <th scope="col">Company</th>
                <th scope="col">Subject</th>
              </tr>
            </thead>
            <tbody>
              {(data?.logs.length ?? 0) === 0 ? (
                <tr>
                  <td colSpan={6}>
                    No notification logs yet. Logs use Training Manager Logs when
                    configured.
                  </td>
                </tr>
              ) : (
                data!.logs.map((entry) => (
                  <tr key={entry.id}>
                    <td>{formatDateTime(entry.createdAt)}</td>
                    <td>{entry.type}</td>
                    <td>{entry.status}</td>
                    <td>{entry.recipientEmail ?? "—"}</td>
                    <td>{entry.companyName ?? "—"}</td>
                    <td>{entry.subject}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>

      <section className={styles.syncPanel}>
        <div className={styles.syncPanelHeader}>
          <strong>Document notification status</strong>
          <span>Recent files</span>
        </div>
        <div className={styles.tableWrap}>
          <table className={styles.dataTable}>
            <thead>
              <tr>
                <th scope="col">Name</th>
                <th scope="col">Company</th>
                <th scope="col">Type</th>
                <th scope="col">Visible</th>
                <th scope="col">Notify</th>
                <th scope="col">Sent</th>
                <th scope="col">Actions</th>
              </tr>
            </thead>
            <tbody>
              {(data?.documentStatus.length ?? 0) === 0 ? (
                <tr>
                  <td colSpan={7}>No documents found.</td>
                </tr>
              ) : (
                data!.documentStatus.map((doc) => (
                  <tr key={doc.id}>
                    <td>{doc.name}</td>
                    <td>{doc.company ?? "—"}</td>
                    <td>{doc.documentType ?? "—"}</td>
                    <td>{doc.customerVisible ? "Yes" : "No"}</td>
                    <td>{doc.notifyCustomer ? "Yes" : "No"}</td>
                    <td>{doc.notificationSent ? "Yes" : "No"}</td>
                    <td>
                      <button
                        type="button"
                        className={styles.linkButton}
                        disabled={busy}
                        onClick={() => void resendDocument(doc.id)}
                      >
                        Send / retry
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

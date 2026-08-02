import "server-only";

import { stripSharePointHtml } from "@/lib/text/stripSharePointHtml";
import type { OutlookEventPayload } from "@/types/calendarSync";

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function toIsoOrFallback(
  value: string | null | undefined,
  fallbackHours = 1,
  base?: string | null,
): string {
  if (value?.trim()) {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }
  }
  const start = base?.trim() ? new Date(base) : new Date();
  if (Number.isNaN(start.getTime())) {
    return new Date().toISOString();
  }
  start.setHours(start.getHours() + fallbackHours);
  return start.toISOString();
}

/**
 * Maps a SharePoint Events row into an Outlook Graph event payload.
 * CustomerVisible does not affect Outlook visibility.
 */
export function mapSharePointEventToOutlookPayload(input: {
  title: string;
  eventDate: string | null;
  endDate: string | null;
  location: string | null;
  description: string | null;
  trainingAddress: string | null;
  companyName: string | null;
  bookingStatus?: "Tentative" | "Confirmed" | null;
}): OutlookEventPayload {
  const description = stripSharePointHtml(input.description) ?? "";
  const trainingAddress = stripSharePointHtml(input.trainingAddress) ?? "";
  const parts: string[] = [];
  if (description) {
    parts.push(`<p>${escapeHtml(description).replace(/\n/g, "<br/>")}</p>`);
  }
  if (input.companyName?.trim()) {
    parts.push(`<p><strong>Company:</strong> ${escapeHtml(input.companyName.trim())}</p>`);
  }
  if (trainingAddress.trim()) {
    parts.push(
      `<p><strong>Training address:</strong> ${escapeHtml(trainingAddress.trim()).replace(/\n/g, "<br/>")}</p>`,
    );
  }
  const statusLabel =
    input.bookingStatus === "Confirmed" ? "Confirmed" : "Tentative (offered)";
  parts.push(
    `<p><strong>Booking status:</strong> ${escapeHtml(statusLabel)}</p>`,
  );
  parts.push(
    `<p><em>Synced from PAVE Training Portal (SharePoint Events).</em></p>`,
  );

  const startIso = toIsoOrFallback(input.eventDate);
  const endIso = toIsoOrFallback(input.endDate, 1, startIso);

  const title = input.title.trim() || "PAVE Training Event";
  const company = input.companyName?.trim() || "";
  const subject = company ? `${company} — ${title}` : title;

  return {
    subject,
    bodyHtml: parts.join("\n"),
    startIso,
    endIso,
    location: input.location?.trim() || null,
    timeZone: "UTC",
    showAs:
      input.bookingStatus === "Confirmed" ? "busy" : "tentative",
  };
}

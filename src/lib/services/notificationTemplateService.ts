import "server-only";

import { readFile } from "fs/promises";
import path from "path";

import { getPortalPublicUrl } from "@/lib/portalUrl";
import { getNotificationSettingsSync } from "@/lib/services/notificationConfig";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { fetchThumbnailContent } from "@/lib/services/listThumbnailService";
import {
  getListItemByKey,
  getListItemsByKey,
} from "@/lib/services/sharePointListService";

export const PAVE_EMAIL_LOGO_CID = "pave-logo";
export const COMPANY_EMAIL_LOGO_CID = "company-logo";

export type InlineEmailAttachment = {
  filename: string;
  contentType: string;
  content: string;
  encoding: "base64";
  contentId: string;
  isInline: true;
};

/**
 * Brand header for emails. When a company logo was attached as
 * `COMPANY_EMAIL_LOGO_CID`, it appears above the PAVE mark.
 */
export function emailLogoHtml(options?: {
  includeCompanyLogo?: boolean;
}): string {
  const companyBlock = options?.includeCompanyLogo
    ? `<img src="cid:${COMPANY_EMAIL_LOGO_CID}" alt="Company logo" width="160" style="display:block;width:160px;max-width:100%;height:auto;border:0;margin:0 0 12px 0;" />`
    : "";
  return `<p style="margin:0 0 20px 0;">
${companyBlock}<img src="cid:${PAVE_EMAIL_LOGO_CID}" alt="PAVE Training" width="140" style="display:block;width:140px;max-width:100%;height:auto;border:0;" />
</p>`;
}

/** Reads public/brand/pave-logo.png for inline email embedding. */
export async function loadPaveLogoAttachment(): Promise<InlineEmailAttachment | null> {
  try {
    const logoPath = path.join(process.cwd(), "public", "brand", "pave-logo.png");
    const bytes = await readFile(logoPath);
    return {
      filename: "pave-logo.png",
      contentType: "image/png",
      content: bytes.toString("base64"),
      encoding: "base64",
      contentId: PAVE_EMAIL_LOGO_CID,
      isInline: true,
    };
  } catch (error) {
    console.warn("[notifications] Could not load PAVE logo for email:", error);
    return null;
  }
}

/**
 * Loads the Company List logo thumbnail as an inline CID attachment.
 * Resolves by SharePoint item id, or by exact company name when id is missing.
 */
export async function loadCompanyLogoAttachment(input: {
  companyId?: string | null;
  companyName?: string | null;
}): Promise<InlineEmailAttachment | null> {
  try {
    const fields = getSharePointFields("company");
    let itemId = String(input.companyId ?? "").trim();
    if (!itemId && input.companyName?.trim()) {
      const nameKey = input.companyName.trim().toLowerCase();
      const rows = await getListItemsByKey("company", { top: 5000 });
      const hit = rows.find((row) => {
        const name =
          String(row.fields[fields.companyName] ?? row.fields.Title ?? "")
            .trim()
            .toLowerCase();
        return name === nameKey;
      });
      itemId = hit?.id ?? "";
    }
    if (!itemId) return null;

    const item = await getListItemByKey("company", itemId);
    if (!item) return null;
    const content = await fetchThumbnailContent(item.fields[fields.companyLogo]);
    if (!content?.bytes) return null;

    const contentType = content.contentType || "image/png";
    const ext = contentType.includes("jpeg")
      ? "jpg"
      : contentType.includes("webp")
        ? "webp"
        : contentType.includes("gif")
          ? "gif"
          : "png";
    return {
      filename: `company-logo.${ext}`,
      contentType,
      content: Buffer.from(content.bytes).toString("base64"),
      encoding: "base64",
      contentId: COMPANY_EMAIL_LOGO_CID,
      isInline: true,
    };
  } catch (error) {
    console.warn("[notifications] Could not load company logo for email:", error);
    return null;
  }
}

/** PAVE + optional company logo attachments for a send. */
export async function loadEmailBrandAttachments(input?: {
  companyId?: string | null;
  companyName?: string | null;
}): Promise<{
  attachments: InlineEmailAttachment[];
  includeCompanyLogo: boolean;
}> {
  const [pave, company] = await Promise.all([
    loadPaveLogoAttachment(),
    input
      ? loadCompanyLogoAttachment(input)
      : Promise.resolve(null),
  ]);
  const attachments = [pave, company].filter(
    (row): row is InlineEmailAttachment => Boolean(row),
  );
  return {
    attachments,
    includeCompanyLogo: Boolean(company),
  };
}

/** Stable portal origin for customer emails — never omit the link. */
export function customerPortalBaseUrl(): string {
  const fromSettings = getNotificationSettingsSync().portalUrl?.trim();
  const base = (fromSettings || getPortalPublicUrl()).replace(/\/+$/, "");
  return base || getPortalPublicUrl();
}

/** Absolute customer-portal URL (path defaults to the external dashboard). */
export function customerPortalUrl(path = "/customer"): string {
  const base = customerPortalBaseUrl();
  const clean = path.startsWith("/") ? path : `/${path}`;
  return `${base}${clean}`;
}

function portalCtaText(
  url: string,
  label = "Open the PAVE Training Portal",
): string {
  return `${label}:\n${url}`;
}

function portalCtaHtml(
  url: string,
  label = "Open the PAVE Training Portal",
): string {
  return `<p><a href="${escapeHtml(url)}">${escapeHtml(label)}</a></p>`;
}

function portalDocumentsUrl(): string {
  return customerPortalUrl("/customer/documents");
}

function portalMatrixUrl(): string {
  return customerPortalUrl("/customer");
}

function portalLoginUrl(): string {
  return customerPortalUrl("/login");
}

function portalEventsUrl(): string {
  return customerPortalUrl("/customer/events");
}

export function documentUploadEmailTemplate(input: {
  companyName: string;
  candidateName?: string | null;
  documentType: string;
  includeCompanyLogo?: boolean;
}): { subject: string; text: string; html: string } {
  const who = input.candidateName?.trim()
    ? `${input.companyName} / ${input.candidateName.trim()}`
    : input.companyName;
  const docsUrl = portalDocumentsUrl();

  const subject = "New training document available";
  const text = [
    `A new document has been shared for ${who}.`,
    `Document type: ${input.documentType}`,
    "",
    "Please log into the PAVE Training Portal to view it.",
    portalCtaText(docsUrl, "Open Documents"),
  ].join("\n");

  const html = `${emailLogoHtml({ includeCompanyLogo: input.includeCompanyLogo })}<p>A new document has been shared for <strong>${escapeHtml(who)}</strong>.</p>
<p>Document type: <strong>${escapeHtml(input.documentType)}</strong></p>
<p>Please log into the PAVE Training Portal to view it.</p>
${portalCtaHtml(docsUrl, "Open Documents in the portal")}`;

  return { subject, text, html };
}

export function expiryReminderEmailTemplate(input: {
  companyName: string;
  windowLabel: string;
  candidateCount: number;
  includeCompanyLogo?: boolean;
}): { subject: string; text: string; html: string } {
  const matrixUrl = portalMatrixUrl();

  const subject = "Training expiry reminder";
  const text = [
    `Training records are expired or due soon for ${input.companyName}.`,
    `Reminder window: ${input.windowLabel}`,
    `Candidates affected: ${input.candidateCount}`,
    "",
    "Please log into the PAVE Training Portal to review the Training Matrix.",
    portalCtaText(matrixUrl, "Open the Training Matrix"),
  ].join("\n");

  const html = `${emailLogoHtml({ includeCompanyLogo: input.includeCompanyLogo })}<p>Training records are expired or due soon for <strong>${escapeHtml(input.companyName)}</strong>.</p>
<p>Reminder window: <strong>${escapeHtml(input.windowLabel)}</strong><br/>
Candidates affected: <strong>${input.candidateCount}</strong></p>
<p>Please log into the PAVE Training Portal to review the Training Matrix.</p>
${portalCtaHtml(matrixUrl, "Open the Training Matrix")}`;

  return { subject, text, html };
}

export function adminAlertEmailTemplate(input: {
  title: string;
  detail: string;
}): { subject: string; text: string; html: string } {
  const adminUrl = customerPortalUrl("/admin");
  const subject = `PAVE Admin alert: ${input.title}`;
  const text = [
    input.title,
    "",
    input.detail,
    "",
    "This alert was generated by the PAVE Training Portal.",
    portalCtaText(adminUrl, "Open the Admin portal"),
  ].join("\n");
  const html = `${emailLogoHtml()}<p><strong>${escapeHtml(input.title)}</strong></p>
<p>${escapeHtml(input.detail)}</p>
<p>This alert was generated by the PAVE Training Portal.</p>
${portalCtaHtml(adminUrl, "Open the Admin portal")}`;
  return { subject, text, html };
}

/**
 * Portal invitation email — used for User / Customer / Training Manager /
 * Supervisor / Candidate invites. Client-approved rule: every invitation MUST
 * include Company and Role on their own lines so the recipient knows exactly
 * which company they are being granted access to and in what capacity.
 * Generic invites without a Company are rejected by `sendPortalInviteNotification`.
 */
export function portalInviteEmailTemplate(input: {
  displayName?: string | null;
  companyName?: string | null;
  roleLabel?: string | null;
  includeCompanyLogo?: boolean;
}): { subject: string; text: string; html: string } {
  const loginUrl = portalLoginUrl();
  const who = input.displayName?.trim() || "there";
  // Fallback labels only appear if a call site slipped past the send-time
  // guard — we still render the field so the recipient can see the drift.
  const company = input.companyName?.trim() || "— not set —";
  const role = input.roleLabel?.trim() || "— not set —";

  const subject = "You're invited to the PAVE Training Portal";
  const text = [
    `Hi ${who},`,
    "",
    "You have been invited to the PAVE Training Portal.",
    "",
    "Company:",
    company,
    "",
    "Role:",
    role,
    "",
    portalCtaText(loginUrl, "Sign in here"),
    "",
    "Use your work email with Microsoft sign-in, or request a one-time code from the login screen. If your organisation uses multi-factor authentication, you will be prompted as usual.",
  ].join("\n");

  const html = `${emailLogoHtml({ includeCompanyLogo: input.includeCompanyLogo })}<p>Hi ${escapeHtml(who)},</p>
<p>You have been invited to the <strong>PAVE Training Portal</strong>.</p>
<p><strong>Company:</strong><br/>${escapeHtml(company)}</p>
<p><strong>Role:</strong><br/>${escapeHtml(role)}</p>
${portalCtaHtml(loginUrl, "Sign in to the portal")}
<p>Use your work email with Microsoft sign-in, or request a one-time code from the login screen. If your organisation uses multi-factor authentication, you will be prompted as usual.</p>`;

  return { subject, text, html };
}

/**
 * Client-approved booking confirmation — sent when Tentative → Confirmed.
 */
export function bookingConfirmedEmailTemplate(input: {
  trainingManagerName?: string | null;
  eventTitle: string;
  dateLabel: string;
  timeLabel: string;
  location?: string | null;
}): { subject: string; text: string; html: string } {
  const eventsUrl = portalEventsUrl();
  const who = input.trainingManagerName?.trim() || "there";
  const title = input.eventTitle.trim() || "Training booking";
  const location = input.location?.trim() || "—";
  const contact = "Info@pavetraining.co.uk";

  const subject = `Training booking confirmed — ${title}`;
  const text = [
    `Hi ${who},`,
    "",
    "Your training booking is confirmed.",
    "",
    `Event: ${title}`,
    `Date: ${input.dateLabel}`,
    `Time: ${input.timeLabel}`,
    `Location: ${location}`,
    "",
    portalCtaText(eventsUrl, "View in the PAVE Training Portal"),
    "",
    "This email includes a calendar invite (.ics). Open / Accept it in Outlook to add the booking to your calendar.",
    "",
    "Any questions please contact:",
    contact,
    "",
    "Cancellation terms:",
    "100% fee if cancelled within 10 days of the course start date.",
    "50% fee if cancelled 11-20 days before the course start date.",
    "These charges will not apply if the course is rebooked for alternative dates.",
    "",
    "Kind regards,",
    "PAVE Training",
  ].join("\n");

  const html = `${emailLogoHtml()}<p>Hi ${escapeHtml(who)},</p>
<p>Your training booking is confirmed.</p>
<p><strong>Event:</strong> ${escapeHtml(title)}<br/>
<strong>Date:</strong> ${escapeHtml(input.dateLabel)}<br/>
<strong>Time:</strong> ${escapeHtml(input.timeLabel)}<br/>
<strong>Location:</strong> ${escapeHtml(location)}</p>
${portalCtaHtml(eventsUrl, "View events in the PAVE Training Portal")}
<p>This email includes a calendar invite (<strong>.ics</strong>). Open or Accept it in Outlook to add the booking to your calendar.</p>
<p>Any questions please contact:<br/>
<a href="mailto:info@pavetraining.co.uk">${escapeHtml(contact)}</a></p>
<p><strong>Cancellation terms:</strong><br/>
100% fee if cancelled within 10 days of the course start date.<br/>
50% fee if cancelled 11-20 days before the course start date.<br/>
These charges will not apply if the course is rebooked for alternative dates.</p>
<p>Kind regards,<br/>PAVE Training</p>`;

  return { subject, text, html };
}

/** Training Manager / customer notify when a register row is added/updated. */
export function trainingRecordChangeEmailTemplate(input: {
  greeting: string;
  registerLabel: string;
  actionLabel: string;
  candidate: string;
  category: string;
  trainingDate: string;
  expiry: string;
  company: string;
  outcome?: string | null;
  actor: string;
  /** NPORS: category title (without N-code). */
  categoryName?: string | null;
  /** NPORS: N001-style code (or card number fallback). */
  nNumber?: string | null;
  includeCompanyLogo?: boolean;
}): { subject: string; text: string; html: string } {
  const portalUrl = portalMatrixUrl();
  const categoryName = input.categoryName?.trim() || input.category;
  const nNumber = input.nNumber?.trim() || null;
  const subjectParts = [
    `Training ${input.actionLabel}: ${input.candidate}`,
    input.registerLabel,
    nNumber || categoryName,
  ];
  const subject = subjectParts.join(" — ");

  const detailLines = [
    `Candidate name: ${input.candidate}`,
    nNumber ? `N number: ${nNumber}` : null,
    `Category name: ${categoryName}`,
    `Training date: ${input.trainingDate}`,
    `Expiry date: ${input.expiry}`,
    `Company: ${input.company}`,
    input.outcome ? `Outcome: ${input.outcome}` : null,
    `Updated by: ${input.actor}`,
  ].filter((line): line is string => line !== null);

  const text = [
    input.greeting,
    "",
    `A ${input.registerLabel} training record has been ${input.actionLabel} in the PAVE Training Portal.`,
    "",
    ...detailLines,
    "",
    "Sign in to the PAVE Training Portal to review the record and any candidate documents.",
    portalCtaText(portalUrl, "Open the portal"),
  ].join("\n");

  const htmlRows = [
    `<tr><td style="padding:2px 12px 2px 0"><strong>Candidate name</strong></td><td style="padding:2px 0">${escapeHtml(input.candidate)}</td></tr>`,
    nNumber
      ? `<tr><td style="padding:2px 12px 2px 0"><strong>N number</strong></td><td style="padding:2px 0">${escapeHtml(nNumber)}</td></tr>`
      : "",
    `<tr><td style="padding:2px 12px 2px 0"><strong>Category name</strong></td><td style="padding:2px 0">${escapeHtml(categoryName)}</td></tr>`,
    `<tr><td style="padding:2px 12px 2px 0"><strong>Training date</strong></td><td style="padding:2px 0">${escapeHtml(input.trainingDate)}</td></tr>`,
    `<tr><td style="padding:2px 12px 2px 0"><strong>Expiry date</strong></td><td style="padding:2px 0">${escapeHtml(input.expiry)}</td></tr>`,
    `<tr><td style="padding:2px 12px 2px 0"><strong>Company</strong></td><td style="padding:2px 0">${escapeHtml(input.company)}</td></tr>`,
    input.outcome
      ? `<tr><td style="padding:2px 12px 2px 0"><strong>Outcome</strong></td><td style="padding:2px 0">${escapeHtml(input.outcome)}</td></tr>`
      : "",
    `<tr><td style="padding:2px 12px 2px 0"><strong>Updated by</strong></td><td style="padding:2px 0">${escapeHtml(input.actor)}</td></tr>`,
  ]
    .filter(Boolean)
    .join("");

  const html = `${emailLogoHtml({ includeCompanyLogo: input.includeCompanyLogo })}<p>${escapeHtml(input.greeting)}</p>
<p>A <strong>${escapeHtml(input.registerLabel)}</strong> training record has been <strong>${escapeHtml(input.actionLabel)}</strong> in the PAVE Training Portal.</p>
<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
  ${htmlRows}
</table>
<p>Sign in to the PAVE Training Portal to review the record and any candidate documents.</p>
${portalCtaHtml(portalUrl, "Open the PAVE Training Portal")}`;

  return { subject, text, html };
}

export function testEmailTemplate(): {
  subject: string;
  text: string;
  html: string;
} {
  const portalUrl = customerPortalBaseUrl();
  const subject = "PAVE Training Portal — test notification";
  const text = [
    "This is a test notification from the PAVE Training Portal.",
    portalCtaText(portalUrl, "Portal"),
    `Sent at: ${new Date().toISOString()}`,
  ].join("\n");
  const html = `${emailLogoHtml()}<p>This is a test notification from the PAVE Training Portal.</p>
${portalCtaHtml(portalUrl, "Open the PAVE Training Portal")}
<p>Sent at: ${escapeHtml(new Date().toISOString())}</p>`;
  return { subject, text, html };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

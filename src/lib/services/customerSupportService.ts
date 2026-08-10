import "server-only";

import { ValidationError } from "@/lib/services/validationService";
import { isValidEmail, normalizeEmail } from "@/lib/validation/email";
import { sendNotification } from "@/lib/services/notificationService";
import type { CustomerContext } from "@/types/models";

const SUPPORT_TO =
  process.env.PAVE_SUPPORT_EMAIL?.trim() || "info@pavetraining.co.uk";

export type CustomerSupportPayload = {
  name: string;
  email: string;
  phone?: string | null;
  companyName?: string | null;
  subject: string;
  message: string;
};

export async function submitCustomerSupportRequest(
  input: CustomerSupportPayload,
  context: CustomerContext,
): Promise<{ ok: true }> {
  const name = input.name?.trim();
  const email = normalizeEmail(input.email);
  const subject = input.subject?.trim();
  const message = input.message?.trim();
  const phone = input.phone?.trim() || null;

  if (!name) throw new ValidationError("Name is required.");
  if (!isValidEmail(email)) {
    throw new ValidationError("A valid email is required.");
  }
  if (!subject) throw new ValidationError("Subject is required.");
  if (!message || message.length < 10) {
    throw new ValidationError("Please enter a message (at least 10 characters).");
  }

  const company =
    input.companyName?.trim() || context.companyName || "Unknown company";

  const mailSubject = `Portal support: ${subject}`;
  const bodyText = [
    "New customer portal support request",
    "",
    `Name: ${name}`,
    `Email: ${email}`,
    `Phone: ${phone ?? "—"}`,
    `Company: ${company}`,
    `Portal role: ${context.roleLabel}`,
    `Signed-in as: ${context.loggedInEmail}`,
    "",
    "Message:",
    message,
  ].join("\n");

  const bodyHtml = `
    <p><strong>New customer portal support request</strong></p>
    <ul>
      <li><strong>Name:</strong> ${escapeHtml(name)}</li>
      <li><strong>Email:</strong> ${escapeHtml(email)}</li>
      <li><strong>Phone:</strong> ${escapeHtml(phone ?? "—")}</li>
      <li><strong>Company:</strong> ${escapeHtml(company)}</li>
      <li><strong>Portal role:</strong> ${escapeHtml(context.roleLabel)}</li>
      <li><strong>Signed-in as:</strong> ${escapeHtml(context.loggedInEmail)}</li>
    </ul>
    <p><strong>Subject:</strong> ${escapeHtml(subject)}</p>
    <p>${escapeHtml(message).replace(/\n/g, "<br/>")}</p>
  `;

  const result = await sendNotification({
    type: "admin_alert",
    to: SUPPORT_TO,
    subject: mailSubject,
    text: bodyText,
    html: bodyHtml,
    companyName: company,
    actorEmail: context.loggedInEmail,
    detail: `Support form from ${email}`,
  });

  if (result.status !== "sent") {
    throw new ValidationError(
      result.errorMessage ||
        "We could not send your message just now. Please email info@pavetraining.co.uk.",
    );
  }

  return { ok: true };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

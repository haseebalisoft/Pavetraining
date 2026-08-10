import "server-only";

import { createHmac, randomInt, timingSafeEqual } from "crypto";

import {
  emailLogoHtml,
  loadPaveLogoAttachment,
} from "@/lib/services/notificationTemplateService";
import { getActivePermissionByEmail } from "@/lib/services/permissionService";
import { sendNotification } from "@/lib/services/notificationService";
import { ValidationError } from "@/lib/services/validationService";
import { isValidEmail, normalizeEmail } from "@/lib/validation/email";

const OTP_TTL_MS = 10 * 60 * 1000;
const OTP_LENGTH = 6;

type OtpChallengePayload = {
  email: string;
  exp: number;
  salt: string;
  hash: string;
};

function authSecret(): string {
  const secret = process.env.AUTH_SECRET?.trim();
  if (!secret) {
    throw new Error("AUTH_SECRET is not configured.");
  }
  return secret;
}

function sha256Hmac(value: string): string {
  return createHmac("sha256", authSecret()).update(value).digest("hex");
}

function generateCode(): string {
  const max = 10 ** OTP_LENGTH;
  return String(randomInt(0, max)).padStart(OTP_LENGTH, "0");
}

function encodeChallenge(payload: OtpChallengePayload): string {
  const body = Buffer.from(JSON.stringify(payload), "utf8").toString(
    "base64url",
  );
  const sig = createHmac("sha256", authSecret()).update(body).digest("base64url");
  return `${body}.${sig}`;
}

function decodeChallenge(token: string): OtpChallengePayload | null {
  const [body, sig] = token.split(".");
  if (!body || !sig) return null;
  const expected = createHmac("sha256", authSecret())
    .update(body)
    .digest("base64url");
  const a = Buffer.from(sig);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !timingSafeEqual(a, b)) return null;
  try {
    const parsed = JSON.parse(
      Buffer.from(body, "base64url").toString("utf8"),
    ) as OtpChallengePayload;
    if (
      typeof parsed.email !== "string" ||
      typeof parsed.exp !== "number" ||
      typeof parsed.salt !== "string" ||
      typeof parsed.hash !== "string"
    ) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

function codeHash(email: string, salt: string, code: string): string {
  return sha256Hmac(`${email}|${salt}|${code}`);
}

/**
 * Sends a one-time code if the email has an Active Permissions row.
 * Always returns a generic success shape for unknown emails (no account enumeration).
 */
export async function requestEmailOtp(input: {
  email: string;
}): Promise<{ ok: true; challenge: string | null; message: string }> {
  const email = normalizeEmail(input.email);
  if (!isValidEmail(email)) {
    throw new ValidationError("Enter a valid email address.");
  }

  const permission = await getActivePermissionByEmail(email);
  if (!permission) {
    return {
      ok: true,
      challenge: null,
      message:
        "If this email is registered, a one-time code has been sent. Check your inbox.",
    };
  }

  const code = generateCode();
  const salt = sha256Hmac(`${email}|${Date.now()}|${randomInt(1, 1_000_000)}`);
  const exp = Date.now() + OTP_TTL_MS;
  const challenge = encodeChallenge({
    email,
    exp,
    salt,
    hash: codeHash(email, salt, code),
  });

  const portalBase = (
    process.env.AUTH_URL?.trim() ||
    process.env.NEXT_PUBLIC_APP_URL?.trim() ||
    (process.env.VERCEL_URL ? `https://${process.env.VERCEL_URL}` : "") ||
    "https://pave-training-portal-nu.vercel.app"
  ).replace(/\/$/, "");

  // Keep the login URL credential-free. Embedding OTP/challenge/query tokens
  // looks like phishing to junk filters (esp. Outlook).
  const loginUrl = `${portalBase}/login`;

  const logo = await loadPaveLogoAttachment();

  const result = await sendNotification({
    type: "login_otp",
    to: email,
    subject: "Your PAVE Training Portal sign-in code",
    text: [
      "Your PAVE Training Portal sign-in code is:",
      "",
      code,
      "",
      "Enter this code on the portal login page:",
      loginUrl,
      "",
      "This code expires in 10 minutes.",
      "If you did not request this, you can ignore this email.",
    ].join("\n"),
    html: `${emailLogoHtml()}<p>Your <strong>PAVE Training Portal</strong> sign-in code is:</p>
<p style="font-size:28px;font-weight:700;letter-spacing:0.18em;font-family:Segoe UI,Arial,sans-serif">${code}</p>
<p>Enter this code on the <a href="${loginUrl}">portal login page</a>.</p>
<p>This code expires in 10 minutes.</p>
<p>If you did not request this, you can ignore this email.</p>`,
    companyName: permission.companyDisplayName ?? null,
    detail: "Email OTP login code",
    fromName: "PAVE Training Portal",
    dedupeKey: `login-otp:${email}:${Math.floor(Date.now() / 60_000)}`,
    attachments: logo ? [logo] : undefined,
  });

  if (
    result.status === "not_configured" ||
    result.status === "failed" ||
    result.status === "skipped"
  ) {
    const raw = (result.errorMessage || "").toLowerCase();
    if (raw.includes("access is denied") || raw.includes("accessdenied")) {
      throw new ValidationError(
        "Email sign-in is temporarily unavailable (mail permission). Please use Sign in with Microsoft, or ask PAVE to enable outbound email.",
      );
    }
    throw new ValidationError(
      result.errorMessage ||
        "Email sign-in is not configured yet. Please use Microsoft sign-in, or ask PAVE to enable outbound email.",
    );
  }

  return {
    ok: true,
    challenge,
    message: "A one-time code has been sent to your email.",
  };
}

/**
 * Validates email + code + challenge for Credentials provider authorize().
 */
export async function verifyEmailOtp(input: {
  email?: string | null;
  code?: string | null;
  challenge?: string | null;
}): Promise<{ email: string }> {
  const email = normalizeEmail(input.email ?? "");
  const code = String(input.code ?? "").trim();
  const challenge = String(input.challenge ?? "").trim();

  if (!isValidEmail(email) || !code || !challenge) {
    throw new ValidationError("Email, code, and challenge are required.");
  }
  if (!/^\d{6}$/.test(code)) {
    throw new ValidationError("Enter the 6-digit code from your email.");
  }

  const payload = decodeChallenge(challenge);
  if (!payload || payload.email !== email) {
    throw new ValidationError("Invalid or expired code. Request a new one.");
  }
  if (Date.now() > payload.exp) {
    throw new ValidationError("This code has expired. Request a new one.");
  }

  const expected = codeHash(email, payload.salt, code);
  const a = Buffer.from(expected);
  const b = Buffer.from(payload.hash);
  if (a.length !== b.length || !timingSafeEqual(a, b)) {
    throw new ValidationError("Invalid code. Check the email and try again.");
  }

  const permission = await getActivePermissionByEmail(email);
  if (!permission) {
    throw new ValidationError(
      "No active portal access for this email. Ask PAVE to add you under Permissions.",
    );
  }

  return { email };
}

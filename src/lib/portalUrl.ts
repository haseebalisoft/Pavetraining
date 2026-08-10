/**
 * Stable public portal origin for emails, OTP, and absolute links.
 * Never prefer VERCEL_URL — that hostname changes on every deployment.
 */
export function getPortalPublicUrl(): string {
  const candidates = [
    process.env.AUTH_URL,
    process.env.NEXT_PUBLIC_APP_URL,
    process.env.VERCEL_PROJECT_PRODUCTION_URL
      ? `https://${process.env.VERCEL_PROJECT_PRODUCTION_URL}`
      : "",
    "https://pave-training-portal-nu.vercel.app",
  ];

  for (const raw of candidates) {
    const value = String(raw ?? "").trim().replace(/\/+$/, "");
    if (!value) continue;
    // Ignore localhost in production deploys so empty/mis-set AUTH_URL
    // does not leak local URLs into live emails.
    if (
      process.env.VERCEL === "1" &&
      /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i.test(value)
    ) {
      continue;
    }
    if (/^https?:\/\//i.test(value)) return value;
  }

  return "https://pave-training-portal-nu.vercel.app";
}

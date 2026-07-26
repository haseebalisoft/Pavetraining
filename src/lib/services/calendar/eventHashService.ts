import "server-only";

import { createHash } from "crypto";

/**
 * Stable hash of the SharePoint event fields that drive Outlook content.
 * Used to skip no-op updates and prepare for future two-way sync.
 */
export function computeEventSyncHash(input: {
  title: string | null | undefined;
  eventDate: string | null | undefined;
  endDate: string | null | undefined;
  location: string | null | undefined;
  description: string | null | undefined;
  trainingAddress: string | null | undefined;
  companyName: string | null | undefined;
  companyId: string | null | undefined;
}): string {
  const normalize = (value: string | null | undefined) =>
    (value ?? "").trim().replace(/\s+/g, " ").toLowerCase();

  const payload = [
    normalize(input.title),
    normalize(input.eventDate),
    normalize(input.endDate),
    normalize(input.location),
    normalize(input.description),
    normalize(input.trainingAddress),
    normalize(input.companyId) || normalize(input.companyName),
  ].join("|");

  return createHash("sha256").update(payload).digest("hex").slice(0, 40);
}

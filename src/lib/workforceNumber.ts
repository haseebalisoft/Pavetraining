/** Portal workforce numbers: W00001, W00002, … (5 digits after W). */
const PORTAL_WORKFORCE_NUMBER = /^W(\d+)$/i;

export function formatPortalWorkforceNumber(sequence: number): string {
  return `W${String(Math.max(1, sequence)).padStart(5, "0")}`;
}

export function parsePortalWorkforceNumberSequence(
  value: string | null | undefined,
): number | null {
  const match = PORTAL_WORKFORCE_NUMBER.exec(value?.trim() ?? "");
  if (!match) return null;
  const parsed = Number.parseInt(match[1], 10);
  return Number.isFinite(parsed) ? parsed : null;
}

/**
 * Next unused portal workforce number. Starts at W00001 when none exist yet.
 * Ignores free-text numbers that are not W#####.
 *
 * Pass previously allocated batch numbers via `extraUsed` so bulk upload of N
 * candidates gets consecutive ids.
 */
export function allocateNextWorkforceNumber(
  workforce: Array<{ workforceNumber: string | null | undefined }>,
  extraUsed: Iterable<string> = [],
): string {
  let max = 0;
  const used = new Set<string>();
  for (const row of workforce) {
    const raw = row.workforceNumber?.trim() ?? "";
    if (raw) used.add(raw.toLowerCase());
    const sequence = parsePortalWorkforceNumberSequence(raw);
    if (sequence !== null && sequence > max) max = sequence;
  }
  for (const value of extraUsed) {
    const raw = value?.trim() ?? "";
    if (raw) used.add(raw.toLowerCase());
    const sequence = parsePortalWorkforceNumberSequence(raw);
    if (sequence !== null && sequence > max) max = sequence;
  }

  let next = max + 1;
  for (;;) {
    const candidate = formatPortalWorkforceNumber(next);
    if (!used.has(candidate.toLowerCase())) return candidate;
    next += 1;
  }
}

/** Portal company numbers: C00001, C00002, … (5 digits after C). */
const PORTAL_COMPANY_NUMBER = /^C(\d+)$/i;

export function formatPortalCompanyNumber(sequence: number): string {
  return `C${String(Math.max(1, sequence)).padStart(5, "0")}`;
}

export function parsePortalCompanyNumberSequence(
  value: string | null | undefined,
): number | null {
  const match = PORTAL_COMPANY_NUMBER.exec(value?.trim() ?? "");
  if (!match) return null;
  const parsed = Number.parseInt(match[1], 10);
  return Number.isFinite(parsed) ? parsed : null;
}

/**
 * Next unused portal company number. Starts at C00001 when none exist yet.
 * Ignores legacy / free-text numbers that are not C##### (e.g. COMP-004).
 *
 * Pass previously allocated batch numbers via `extraUsed` so bulk upload of N
 * rows gets consecutive ids (e.g. C00003 … C00012).
 */
export function allocateNextCompanyNumber(
  companies: Array<{ companyNumber: string | null | undefined }>,
  extraUsed: Iterable<string> = [],
): string {
  let max = 0;
  const used = new Set<string>();
  for (const company of companies) {
    const raw = company.companyNumber?.trim() ?? "";
    if (raw) used.add(raw.toLowerCase());
    const sequence = parsePortalCompanyNumberSequence(raw);
    if (sequence !== null && sequence > max) max = sequence;
  }
  for (const value of extraUsed) {
    const raw = value?.trim() ?? "";
    if (raw) used.add(raw.toLowerCase());
    const sequence = parsePortalCompanyNumberSequence(raw);
    if (sequence !== null && sequence > max) max = sequence;
  }

  let next = max + 1;
  for (;;) {
    const candidate = formatPortalCompanyNumber(next);
    if (!used.has(candidate.toLowerCase())) return candidate;
    next += 1;
  }
}

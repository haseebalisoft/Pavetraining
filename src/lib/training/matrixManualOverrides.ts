/**
 * Manual matrix date overrides — stored as pipe-separated headers in
 * Training Matrix Update column `ManualOverrides` (Text).
 * Sync will not overwrite these headers from registers.
 */

const MANUAL_FIELD = "ManualOverrides";

export function parseManualOverrides(raw: unknown): string[] {
  if (raw == null) return [];
  const text = String(raw).trim();
  if (!text) return [];
  return Array.from(
    new Set(
      text
        .split("|")
        .map((part) => part.trim())
        .filter(Boolean),
    ),
  );
}

export function serializeManualOverrides(headers: string[]): string {
  return Array.from(new Set(headers.map((h) => h.trim()).filter(Boolean))).join(
    "|",
  );
}

export function mergeManualOverrides(
  existing: string[],
  added: string[],
): string[] {
  return Array.from(
    new Set(
      [...existing, ...added].map((h) => h.trim()).filter(Boolean),
    ),
  );
}

export function isManualOverrideHeader(
  header: string,
  overrides: string[] | null | undefined,
): boolean {
  if (!overrides?.length) return false;
  const key = header.trim().toLowerCase();
  return overrides.some((h) => h.trim().toLowerCase() === key);
}

export function isAsbestosAwarenessCategory(
  category: string | null | undefined,
): boolean {
  const text = (category ?? "").trim().toLowerCase();
  if (!text) return false;
  return (
    text.includes("asbestos") ||
    text === "n031" ||
    text.startsWith("n031 ") ||
    text.startsWith("n031-") ||
    text.startsWith("n031 –") ||
    text.startsWith("n031 —")
  );
}

export const ASBESTOS_MATRIX_HEADER = "N031 - Asbestos Awareness";
export const MANUAL_OVERRIDES_FIELD = MANUAL_FIELD;

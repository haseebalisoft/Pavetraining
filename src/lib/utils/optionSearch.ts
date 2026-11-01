export interface SearchableOption {
  value: string;
  label: string;
}

/**
 * Category codes are typed loosely ("n17", "17", "N017", "702a") — turn
 * them into the stored N### form so they match the option value.
 */
function asCategoryCode(query: string): string | null {
  const match = /^n?\s*0*(\d{1,3})([a-z]?)$/i.exec(query.trim());
  if (!match) return null;
  return `N${match[1].padStart(3, "0")}${match[2].toUpperCase()}`;
}

/**
 * Filter options by name or code, best match first:
 *   1. exact code ("n17" → N017)
 *   2. code starts with what was typed ("N70" → N702A, N703 …)
 *   3. every typed word appears in the label ("excav 360" → N202)
 * An empty query returns the options unchanged.
 */
export function searchOptions<T extends SearchableOption>(
  options: T[],
  query: string,
): T[] {
  const trimmed = query.trim().toLowerCase();
  if (!trimmed) return options;

  const code = asCategoryCode(trimmed);
  const words = trimmed.split(/\s+/).filter(Boolean);
  const scored: Array<{ option: T; score: number; index: number }> = [];

  options.forEach((option, index) => {
    const value = option.value.trim().toLowerCase();
    const label = option.label.trim().toLowerCase();
    let score = 0;
    if (code && value === code.toLowerCase()) score = 4;
    else if (value === trimmed) score = 4;
    else if (value.startsWith(trimmed.replace(/\s+/g, ""))) score = 3;
    else if (label.startsWith(trimmed)) score = 2;
    else if (words.every((word) => label.includes(word) || value.includes(word))) {
      score = 1;
    }
    if (score) scored.push({ option, score, index });
  });

  return scored
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .map((entry) => entry.option);
}

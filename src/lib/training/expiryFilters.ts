export type ExpiryFilter =
  | "all"
  | "expired"
  | "expiring-3m"
  | "expiring-6m"
  | "expiring-9m"
  | "missing";

export type ExpiryTone = "missing" | "expired" | "critical" | "warning" | "ok";

const MS_PER_DAY = 86_400_000;

export function daysUntilExpiry(
  expiry: string | null | undefined,
  now = new Date(),
): number | null {
  if (!expiry?.trim()) {
    return null;
  }

  const expiryDate = new Date(expiry);
  if (Number.isNaN(expiryDate.getTime())) {
    return null;
  }

  const startOfToday = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
  );
  const startOfExpiry = new Date(
    expiryDate.getFullYear(),
    expiryDate.getMonth(),
    expiryDate.getDate(),
  );

  return Math.ceil(
    (startOfExpiry.getTime() - startOfToday.getTime()) / MS_PER_DAY,
  );
}

export function getExpiryTone(
  expiry: string | null | undefined,
  now = new Date(),
): ExpiryTone {
  const days = daysUntilExpiry(expiry, now);
  if (days === null) return "missing";
  if (days < 0) return "expired";
  if (days <= 90) return "critical";
  if (days <= 180) return "warning";
  return "ok";
}

export function matchesExpiryFilter(
  expiry: string | null | undefined,
  filter: ExpiryFilter,
  now = new Date(),
): boolean {
  if (filter === "all") {
    return true;
  }

  const days = daysUntilExpiry(expiry, now);

  switch (filter) {
    case "missing":
      return days === null;
    case "expired":
      return days !== null && days < 0;
    case "expiring-3m":
      return days !== null && days >= 0 && days <= 90;
    case "expiring-6m":
      return days !== null && days >= 0 && days <= 180;
    case "expiring-9m":
      return days !== null && days >= 0 && days <= 270;
    default:
      return true;
  }
}

export function formatDisplayDate(value: string | null | undefined): string {
  if (!value?.trim()) {
    return "—";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(date);
}

import { redirect } from "next/navigation";

export const dynamic = "force-dynamic";

/**
 * Training Matrix lives on the customer landing page (`/customer`).
 * Preserve query filters when users hit the old route.
 */
export default async function CustomerTrainingMatrixPage({
  searchParams,
}: {
  searchParams: Promise<Record<string, string | string[] | undefined>>;
}) {
  const params = await searchParams;
  const query = new URLSearchParams();
  for (const [key, value] of Object.entries(params)) {
    if (value === undefined) continue;
    if (Array.isArray(value)) {
      for (const entry of value) query.append(key, entry);
    } else {
      query.set(key, value);
    }
  }
  const qs = query.toString();
  redirect(qs ? `/customer?${qs}` : "/customer");
}

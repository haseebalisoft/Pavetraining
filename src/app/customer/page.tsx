import { redirect } from "next/navigation";

export const dynamic = "force-dynamic";

/**
 * Customer portal landing — always the Dashboard.
 *
 * Client rule (spec):
 *   Admin        → /admin
 *   Customer     → /customer/dashboard
 *   Training Mgr → /customer/dashboard
 *   Supervisor   → /customer/dashboard
 *   Candidate    → /customer/dashboard
 *
 * Customers must NEVER land on the Training Matrix, a candidate profile, or
 * a blank page. The Training Matrix has its own route at
 * `/customer/training-matrix`; this file only exists to keep the bare
 * `/customer` URL working for links and email history.
 */
export default function CustomerHomePage() {
  redirect("/customer/dashboard");
}

import { redirect } from "next/navigation";

export const dynamic = "force-dynamic";

/** Training Matrix is the customer home (`/customer`). */
export default function CustomerTrainingMatrixRedirectPage() {
  redirect("/customer");
}

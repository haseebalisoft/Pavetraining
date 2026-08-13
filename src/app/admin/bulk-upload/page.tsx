import { redirect } from "next/navigation";

import { AdminBulkUploadClient } from "@/components/admin/pages/AdminBulkUploadClient";
import { requireAdminAccess } from "@/lib/services/securityService";

export const dynamic = "force-dynamic";

export default async function AdminBulkUploadPage() {
  // Defence-in-depth for requirement "Use the server-side resolved role,
  // not browser state" — even if the nav is bypassed, Training Managers
  // land back on /admin instead of seeing this page.
  const context = await requireAdminAccess();
  if (!context.isSharePointAdmin) {
    redirect("/admin");
  }
  return <AdminBulkUploadClient />;
}

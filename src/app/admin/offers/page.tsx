import { AdminOffersClient } from "@/components/admin/pages/AdminOffersClient";
import { listAdminOffers } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminOffersPage() {
  const records = await listAdminOffers();
  return <AdminOffersClient initialRows={records} />;
}

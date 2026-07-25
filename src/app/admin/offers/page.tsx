import { AdminOffersClient } from "@/components/admin/pages/AdminOffersClient";
import {
  listAdminCompanies,
  listAdminOffers,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminOffersPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminOffers(),
  ]);
  return <AdminOffersClient companies={companies} initialRows={records} />;
}

import { AdminEventsClient } from "@/components/admin/pages/AdminEventsClient";
import {
  listAdminCompanies,
  listAdminEvents,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminEventsPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminEvents(),
  ]);
  return <AdminEventsClient companies={companies} initialRows={records} />;
}

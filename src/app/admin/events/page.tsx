import { AdminEventsClient } from "@/components/admin/pages/AdminEventsClient";
import {
  getEventsSchemaWarnings,
  listAdminCompanies,
  listAdminEvents,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminEventsPage() {
  const [companies, records, warnings] = await Promise.all([
    listAdminCompanies(),
    listAdminEvents(),
    getEventsSchemaWarnings(),
  ]);
  return (
    <AdminEventsClient
      companies={companies}
      initialRows={records}
      warnings={warnings}
    />
  );
}

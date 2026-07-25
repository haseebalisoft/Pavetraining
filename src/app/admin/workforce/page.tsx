import { AdminWorkforceClient } from "@/components/admin/pages/AdminWorkforceClient";
import {
  listAdminCompanies,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminWorkforcePage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  return (
    <AdminWorkforceClient companies={companies} initialRows={records} />
  );
}

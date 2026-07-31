import { AdminWorkforceClient } from "@/components/admin/pages/AdminWorkforceClient";
import {
  listAdminCompanies,
  listAdminPermissions,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminWorkforcePage() {
  const [companies, records, permissionPeople] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
    listAdminPermissions(),
  ]);
  return (
    <AdminWorkforceClient
      companies={companies}
      initialRows={records}
      permissionPeople={permissionPeople}
    />
  );
}

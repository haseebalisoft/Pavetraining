import { AdminWorkforceClient } from "@/components/admin/pages/AdminWorkforceClient";
import {
  listAdminCompanies,
  listAdminPermissions,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";
import { listAdminDepartments } from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

export default async function AdminWorkforcePage() {
  const [companies, records, permissionPeople, departments] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
    listAdminPermissions(),
    listAdminDepartments(),
  ]);
  return (
    <AdminWorkforceClient
      companies={companies}
      departments={departments}
      initialRows={records}
      permissionPeople={permissionPeople}
    />
  );
}

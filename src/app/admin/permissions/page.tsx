import { AdminPermissionsClient } from "@/components/admin/pages/AdminPermissionsClient";
import {
  listAdminCompanies,
  listAdminPermissions,
} from "@/lib/services/adminCrudService";
import { listAdminDepartments } from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

export default async function AdminPermissionsPage() {
  const [companies, records, departments] = await Promise.all([
    listAdminCompanies(),
    listAdminPermissions(),
    listAdminDepartments(),
  ]);
  return (
    <AdminPermissionsClient
      companies={companies}
      departments={departments}
      initialRows={records}
    />
  );
}

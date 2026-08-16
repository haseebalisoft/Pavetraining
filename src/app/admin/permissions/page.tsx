import { redirect } from "next/navigation";

import { AdminPermissionsClient } from "@/components/admin/pages/AdminPermissionsClient";
import {
  listAdminCompanies,
  listAdminPermissions,
} from "@/lib/services/adminCrudService";
import {
  isDepartmentActive,
  listAdminDepartments,
} from "@/lib/services/departmentService";
import { requireAdminAccess } from "@/lib/services/securityService";

export const dynamic = "force-dynamic";

export default async function AdminPermissionsPage() {
  // Permissions can grant/revoke portal access — pure SharePoint Admins only.
  const context = await requireAdminAccess();
  if (!context.isSharePointAdmin) {
    redirect("/admin");
  }

  const [companies, records, departments] = await Promise.all([
    listAdminCompanies(),
    listAdminPermissions(),
    listAdminDepartments(),
  ]);
  return (
    <AdminPermissionsClient
      companies={companies}
      departments={departments.filter(isDepartmentActive)}
      initialRows={records}
    />
  );
}

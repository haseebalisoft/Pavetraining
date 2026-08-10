import { AdminDepartmentsClient } from "@/components/admin/pages/AdminDepartmentsClient";
import { listAdminCompanies } from "@/lib/services/adminCrudService";
import { listAdminDepartments } from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

export default async function AdminDepartmentsPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    // Admin screen shows Inactive too, so they can be reviewed and reactivated.
    listAdminDepartments(null, { includeInactive: true }),
  ]);
  return (
    <AdminDepartmentsClient companies={companies} initialRows={records} />
  );
}

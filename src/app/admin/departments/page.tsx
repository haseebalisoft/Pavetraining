import { AdminDepartmentsClient } from "@/components/admin/pages/AdminDepartmentsClient";
import { listAdminCompanies } from "@/lib/services/adminCrudService";
import { listAdminDepartments } from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

export default async function AdminDepartmentsPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminDepartments(),
  ]);
  return (
    <AdminDepartmentsClient companies={companies} initialRows={records} />
  );
}

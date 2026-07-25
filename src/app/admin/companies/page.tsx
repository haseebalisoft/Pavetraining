import { AdminCompaniesClient } from "@/components/admin/pages/AdminCompaniesClient";
import { listAdminCompanies } from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminCompaniesPage() {
  const companies = await listAdminCompanies();
  return <AdminCompaniesClient initialRows={companies} />;
}

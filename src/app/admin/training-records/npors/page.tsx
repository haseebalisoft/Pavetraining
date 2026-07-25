import { AdminRegisterClient } from "@/components/admin/pages/AdminRegisterClient";
import {
  listAdminCompanies,
  listAdminRegister,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminNporsPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminRegister("nporsRegister"),
  ]);
  return (
    <AdminRegisterClient
      kind="npors"
      companies={companies}
      initialRows={records}
    />
  );
}

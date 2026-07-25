import { AdminRegisterClient } from "@/components/admin/pages/AdminRegisterClient";
import {
  listAdminCompanies,
  listAdminRegister,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminEusrPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminRegister("eusrRegister"),
  ]);
  return (
    <AdminRegisterClient
      kind="eusr"
      companies={companies}
      initialRows={records}
    />
  );
}

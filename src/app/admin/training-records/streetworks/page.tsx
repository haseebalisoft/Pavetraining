import { AdminRegisterClient } from "@/components/admin/pages/AdminRegisterClient";
import {
  listAdminCompanies,
  listAdminRegister,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminStreetworksPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminRegister("nrswaRegister"),
  ]);
  return (
    <AdminRegisterClient
      kind="streetworks"
      companies={companies}
      initialRows={records}
    />
  );
}

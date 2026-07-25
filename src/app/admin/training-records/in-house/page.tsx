import { AdminRegisterClient } from "@/components/admin/pages/AdminRegisterClient";
import {
  listAdminCompanies,
  listAdminRegister,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminInHousePage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminRegister("inHouseCertificates"),
  ]);
  return (
    <AdminRegisterClient
      kind="in-house"
      companies={companies}
      initialRows={records}
    />
  );
}

import { AdminPermissionsClient } from "@/components/admin/pages/AdminPermissionsClient";
import {
  listAdminCompanies,
  listAdminPermissions,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminPermissionsPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminPermissions(),
  ]);
  return (
    <AdminPermissionsClient companies={companies} initialRows={records} />
  );
}

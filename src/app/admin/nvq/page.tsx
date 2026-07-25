import { AdminNvqClient } from "@/components/admin/pages/AdminNvqClient";
import {
  listAdminCompanies,
  listAdminNvq,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminNvqPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminNvq(),
  ]);
  return <AdminNvqClient companies={companies} initialRows={records} />;
}

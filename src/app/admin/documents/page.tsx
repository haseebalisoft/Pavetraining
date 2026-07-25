import { AdminDocumentsClient } from "@/components/admin/pages/AdminDocumentsClient";
import {
  listAdminCompanies,
  listAdminDocuments,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminDocumentsPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminDocuments(),
  ]);
  return <AdminDocumentsClient companies={companies} initialRows={records} />;
}

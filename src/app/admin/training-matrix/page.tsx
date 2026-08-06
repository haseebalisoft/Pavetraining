import { AdminMatrixClient } from "@/components/admin/pages/AdminMatrixClient";
import {
  listAdminCompanies,
  listAdminMatrix,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminTrainingMatrixPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    // Admin view loads orphan/needs-review rows too; the client hides them
    // behind a "Show all" toggle. (Customer/default callers stay orphan-free.)
    listAdminMatrix(null, { includeUnlinked: true }),
  ]);
  return <AdminMatrixClient companies={companies} initialRows={records} />;
}

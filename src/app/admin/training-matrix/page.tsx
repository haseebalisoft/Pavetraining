import { AdminMatrixClient } from "@/components/admin/pages/AdminMatrixClient";
import {
  listAdminCompanies,
  listAdminMatrix,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminTrainingMatrixPage() {
  const [companies, records] = await Promise.all([
    listAdminCompanies(),
    listAdminMatrix(),
  ]);
  return <AdminMatrixClient companies={companies} initialRows={records} />;
}

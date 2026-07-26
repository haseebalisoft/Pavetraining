import { AdminDocumentsClient } from "@/components/admin/pages/AdminDocumentsClient";
import {
  listAdminCompanies,
  listAdminDocumentsAtPath,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminDocumentsPage() {
  const [companies, records, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminDocumentsAtPath([]),
    listAdminWorkforce(),
  ]);
  return (
    <AdminDocumentsClient
      companies={companies}
      initialRows={records}
      initialPath={[]}
      initialWorkforce={workforce.map((row) => ({
        id: row.id,
        candidateName: row.candidateName,
        companyName: row.companyName,
        workforceNumber: row.workforceNumber,
      }))}
    />
  );
}

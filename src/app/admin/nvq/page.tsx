import { AdminNvqClient } from "@/components/admin/pages/AdminNvqClient";
import {
  listAdminCompanies,
  listAdminNvq,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminNvqPage() {
  const [companies, workforce, records] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
    listAdminNvq(),
  ]);
  return (
    <AdminNvqClient
      companies={companies}
      workforce={workforce.map((row) => ({
        id: row.id,
        candidateName: row.candidateName,
        companyId: row.companyId,
        companyName: row.companyName,
        nporsNumbers: row.nporsNumbers,
        eusrNumber: row.eusrNumber,
        swqrNumber: row.swqrNumber,
        inHouseCertificationNumber: row.inHouseCertificationNumber,
        workforceNumber: row.workforceNumber,
        niNumber: row.niNumber,
      }))}
      initialRows={records}
    />
  );
}

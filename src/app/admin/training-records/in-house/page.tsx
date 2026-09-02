import { AdminRegisterClient } from "@/components/admin/pages/AdminRegisterClient";
import {
  listAdminCompanies,
  listAdminRegister,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";

export const dynamic = "force-dynamic";

export default async function AdminInHousePage() {
  const [companies, workforce, records] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
    listAdminRegister("inHouseCertificates"),
  ]);
  return (
    <AdminRegisterClient
      kind="in-house"
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

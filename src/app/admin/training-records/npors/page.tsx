import { AdminRegisterClient } from "@/components/admin/pages/AdminRegisterClient";
import {
  listAdminCompanies,
  listAdminRegister,
  listAdminWorkforce,
} from "@/lib/services/adminCrudService";
import { listAdminNporsCategoryOptions } from "@/lib/services/nporsCategoriesService";

export const dynamic = "force-dynamic";

export default async function AdminNporsPage() {
  const [companies, workforce, records, nporsCategoryOptions] =
    await Promise.all([
      listAdminCompanies(),
      listAdminWorkforce(),
      listAdminRegister("nporsRegister"),
      listAdminNporsCategoryOptions(),
    ]);
  return (
    <AdminRegisterClient
      kind="npors"
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
      nporsCategoryOptions={nporsCategoryOptions}
    />
  );
}

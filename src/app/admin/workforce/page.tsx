import { AdminWorkforceClient } from "@/components/admin/pages/AdminWorkforceClient";
import {
  listAdminCompanies,
  listAdminPermissions,
  listAdminWorkforce,
  type AdminPermissionRecord,
} from "@/lib/services/adminCrudService";
import {
  isDepartmentActive,
  listAdminDepartments,
} from "@/lib/services/departmentService";

export const dynamic = "force-dynamic";

function permissionPeopleFromRecords(rows: AdminPermissionRecord[]) {
  return rows.map((row) => ({
    id: row.id,
    name: row.name,
    userEmail: row.userEmail,
    roleType: row.roleType,
    companyId: row.companyId,
    status: row.status,
  }));
}

export default async function AdminWorkforcePage() {
  // Fetch shared lists once — listAdminWorkforce used to re-fetch companies /
  // permissions / departments on top of the page's own Promise.all (very slow
  // after a large bulk import).
  const companies = await listAdminCompanies();
  const [departments, permissionPeople] = await Promise.all([
    listAdminDepartments(),
    listAdminPermissions(companies),
  ]);
  const records = await listAdminWorkforce(null, {
    companies,
    people: permissionPeopleFromRecords(permissionPeople),
    departments: departments.map((row) => ({ id: row.id, name: row.name })),
  });

  return (
    <AdminWorkforceClient
      companies={companies}
      // Dropdown only offers Active departments; an existing candidate whose
      // department was later deactivated still displays correctly (name
      // resolution above uses the unfiltered `departments` list).
      departments={departments.filter(isDepartmentActive)}
      initialRows={records}
      permissionPeople={permissionPeople}
    />
  );
}

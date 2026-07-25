"use client";

import { useRouter } from "next/navigation";

import type { Company } from "@/types/models";

import styles from "./admin.module.css";

interface AdminCompanySelectorProps {
  companies: Company[];
  selectedCompanyId: string | null;
}

export function AdminCompanySelector({
  companies,
  selectedCompanyId,
}: AdminCompanySelectorProps) {
  const router = useRouter();

  return (
    <label className={styles.field}>
      <span className={styles.fieldLabel}>Company filter</span>
      <select
        className={styles.select}
        value={selectedCompanyId ?? ""}
        onChange={(event) => {
          const value = event.target.value;
          const url = value
            ? `/admin?companyId=${encodeURIComponent(value)}`
            : "/admin";
          router.push(url);
        }}
      >
        <option value="">All companies</option>
        {companies.map((company) => (
          <option key={company.id} value={company.id}>
            {company.companyName}
            {company.status ? ` (${company.status})` : ""}
          </option>
        ))}
      </select>
    </label>
  );
}

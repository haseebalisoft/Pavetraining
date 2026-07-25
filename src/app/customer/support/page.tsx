import { CustomerPageHeader } from "@/components/customer/CustomerPageHeader";
import styles from "@/components/customer/customer.module.css";

export const dynamic = "force-dynamic";

export default function CustomerSupportPage() {
  return (
    <div>
      <CustomerPageHeader
        title="Support"
        subtitle="Need help with training records, documents, or portal access? Contact your PAVE training manager."
        breadcrumbs={[
          { label: "Customer", href: "/customer" },
          { label: "Support" },
        ]}
      />
      <section className={styles.supportCard}>
        <h2 className={styles.supportTitle}>How to get help</h2>
        <p className={styles.supportCopy}>
          For training renewals, certificate questions, or document requests,
          contact your assigned Training Manager. For portal sign-in problems,
          include the email address you use to sign in.
        </p>
        <ul className={styles.supportList}>
          <li>Have your company name ready</li>
          <li>Mention the candidate or document if relevant</li>
          <li>Do not share passwords or Microsoft account codes</li>
        </ul>
      </section>
    </div>
  );
}

import { BrandLogo } from "@/components/brand/BrandLogo";
import styles from "./offline.module.css";

export const metadata = {
  title: "Offline · PAVE Training Portal",
  description: "You are offline. Reconnect to use the PAVE Training Portal.",
};

export default function OfflinePage() {
  return (
    <main className={styles.page}>
      <BrandLogo variant="compact" />
      <h1 className={styles.title}>You&apos;re offline</h1>
      <p className={styles.copy}>
        The PAVE Training Portal needs a connection for login and SharePoint
        data. Check your network, then try again.
      </p>
      <a className={styles.button} href="/">
        Retry
      </a>
    </main>
  );
}

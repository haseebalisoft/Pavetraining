import styles from "./ui.module.css";

export function EmptyState({
  title,
  description,
}: {
  title: string;
  description: string;
}) {
  return (
    <div className={styles.emptyState}>
      <h2>{title}</h2>
      <p>{description}</p>
    </div>
  );
}

export function LoadingState({ label = "Loading…" }: { label?: string }) {
  return (
    <div className={styles.loadingState} role="status" aria-live="polite">
      <span className={styles.spinner} aria-hidden />
      <p>{label}</p>
    </div>
  );
}

export function ComingNextNotice({
  title,
  description,
}: {
  title: string;
  description: string;
}) {
  return (
    <div className={styles.comingNext}>
      <p className={styles.comingBadge}>Coming next</p>
      <h1>{title}</h1>
      <p>{description}</p>
    </div>
  );
}

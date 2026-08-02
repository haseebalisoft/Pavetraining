import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import type { TrainingTrackerStats } from "../../models";
import { StatsService } from "../../services/StatsService";
import { paveTheme } from "../../theme/paveTheme";
import styles from "./TrainingTracker.module.scss";

export interface ITrainingTrackerProps {
  context: WebPartContext;
}

const LABELS: Array<{ key: keyof TrainingTrackerStats; label: string }> = [
  { key: "totalOperators", label: "Active operators" },
  { key: "expiringIn30Days", label: "Expiring in 30 days" },
  { key: "activeRegistrations", label: "Active registrations" },
  { key: "completedThisMonth", label: "Completed this month" },
];

export const TrainingTracker: React.FC<ITrainingTrackerProps> = (props) => {
  const { context } = props;
  const [stats, setStats] = React.useState<TrainingTrackerStats | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [failed, setFailed] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      setFailed(false);
      try {
        const result = await StatsService.getTrainingTrackerStats(context);
        if (!cancelled) {
          setStats(result);
          setFailed(false);
        }
      } catch (error) {
        console.warn("[TrainingTracker] Unable to load stats", error);
        if (!cancelled) {
          setStats(null);
          setFailed(true);
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [context]);

  return (
    <section
      className={styles.band}
      aria-label="Training tracker"
      style={
        {
          ["--pave-green" as string]: paveTheme.green,
          ["--pave-charcoal-dark" as string]: paveTheme.charcoalDark,
          ["--pave-font-body" as string]: paveTheme.fontBody,
        } as React.CSSProperties
      }
    >
      {failed ? (
        <p className={styles.loadNote}>Unable to load — check permissions</p>
      ) : (
        <div
          className={styles.grid}
          aria-busy={loading}
          aria-label={loading ? "Loading training stats" : "Training stats"}
        >
          {LABELS.map((item) => (
            <div key={item.key} className={styles.stat}>
              <p className={styles.value}>
                {loading || !stats ? (
                  <span className={styles.skeletonValue} aria-hidden="true" />
                ) : (
                  <span aria-label={`${item.label}: ${stats[item.key]}`}>
                    {stats[item.key]}
                  </span>
                )}
              </p>
              <p className={styles.label}>{item.label}</p>
            </div>
          ))}
        </div>
      )}
    </section>
  );
};

export default TrainingTracker;

import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import type { PopularDocument } from "../../models";
import { DocumentsService } from "../../services/DocumentsService";
import { paveTheme } from "../../theme/paveTheme";
import styles from "./PopularDocs.module.scss";

export interface IPopularDocsProps {
  context: WebPartContext;
  documentsUrl: string;
}

function relativeTime(iso: string | null): string {
  if (!iso) return "Edited recently";
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "Edited recently";

  const seconds = Math.round((Date.now() - date.getTime()) / 1000);
  if (seconds < 60) return "Edited just now";
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) return `Edited ${minutes}m ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `Edited ${hours}h ago`;
  const days = Math.round(hours / 24);
  if (days < 30) return `Edited ${days}d ago`;
  const months = Math.round(days / 30);
  if (months < 12) return `Edited ${months}mo ago`;
  return `Edited ${Math.round(months / 12)}y ago`;
}

function iconClass(iconType: string, fileName: string): string {
  const ext = (fileName.split(".").pop() || "").toLowerCase();
  if (iconType === "word" || ext === "doc" || ext === "docx") return styles.iconWord;
  if (iconType === "excel" || ext === "xls" || ext === "xlsx" || ext === "csv") {
    return styles.iconExcel;
  }
  if (iconType === "pdf" || ext === "pdf") return styles.iconPdf;
  if (iconType === "image") return styles.iconImage;
  return styles.iconDefault;
}

export const PopularDocs: React.FC<IPopularDocsProps> = (props) => {
  const { context, documentsUrl } = props;
  const [docs, setDocs] = React.useState<PopularDocument[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [failed, setFailed] = React.useState(false);

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      setFailed(false);
      try {
        const rows = await DocumentsService.getPopular(context, 6);
        if (!cancelled) {
          setDocs(rows);
          setFailed(false);
        }
      } catch (error) {
        console.warn("[PopularDocs] Unable to load documents", error);
        if (!cancelled) {
          setDocs([]);
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

  const seeAll = (documentsUrl || "").trim();

  return (
    <section
      className={styles.section}
      aria-label="Popular documents"
      style={
        {
          ["--pave-green" as string]: paveTheme.green,
          ["--pave-charcoal-dark" as string]: paveTheme.charcoalDark,
          ["--pave-font-body" as string]: paveTheme.fontBody,
        } as React.CSSProperties
      }
    >
      <div className={styles.header}>
        <h2 className={styles.title}>Popular documents</h2>
        {seeAll ? (
          <a
            className={styles.seeAll}
            href={seeAll}
            aria-label="See all popular documents"
          >
            See all
          </a>
        ) : (
          <span className={styles.seeAllMuted}>See all</span>
        )}
      </div>

      {loading ? (
        <div className={styles.grid} aria-busy="true" aria-label="Loading documents">
          {[0, 1, 2, 3, 4, 5].map((key) => (
            <div key={key} className={`${styles.card} ${styles.skeleton}`} />
          ))}
        </div>
      ) : failed ? (
        <p className={styles.loadNote}>Unable to load — check permissions</p>
      ) : docs.length === 0 ? (
        <p className={styles.empty}>No popular documents yet.</p>
      ) : (
        <div className={styles.grid}>
          {docs.map((doc) => (
            <a
              key={doc.id}
              className={styles.card}
              href={doc.url || seeAll || "#"}
              target="_blank"
              rel="noreferrer"
              aria-label={`Open document ${doc.name}`}
            >
              <span
                className={`${styles.fileIcon} ${iconClass(doc.iconType, doc.name)}`}
                aria-hidden="true"
              />
              <span className={styles.fileName}>{doc.name}</span>
              <span className={styles.caption}>{relativeTime(doc.modified)}</span>
            </a>
          ))}
        </div>
      )}
    </section>
  );
};

export default PopularDocs;

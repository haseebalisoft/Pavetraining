import * as React from "react";
import {
  Certificate24Regular,
  Folder24Regular,
  LockClosed24Regular,
  People24Regular,
} from "@fluentui/react-icons";

import { paveTheme } from "../../theme/paveTheme";
import styles from "./QuickTiles.module.scss";

export interface IQuickTilesProps {
  workforceUrl: string;
  eusrUrl: string;
  customerDocsUrl: string;
  adminDocsUrl: string;
  workforceImageUrl?: string;
  eusrImageUrl?: string;
  customerDocsImageUrl?: string;
  adminDocsImageUrl?: string;
}

type TileTone = "green" | "charcoal";

interface ITile {
  id: string;
  title: string;
  description: string;
  url: string;
  tone: TileTone;
  imageUrl?: string;
  icon: React.ReactNode;
}

function openUrl(url: string, label: string): void {
  const target = (url || "").trim();
  if (!target) {
    console.warn(
      `[QuickTiles] ${label} URL is not configured in the property pane`
    );
    return;
  }
  window.location.assign(target);
}

export const QuickTiles: React.FC<IQuickTilesProps> = (props) => {
  const tiles: ITile[] = [
    {
      id: "workforce",
      title: "WorkForce",
      description: "Operators and staff",
      url: props.workforceUrl,
      tone: "green",
      imageUrl: props.workforceImageUrl,
      icon: <People24Regular style={{ width: 40, height: 40 }} aria-hidden />,
    },
    {
      id: "eusr",
      title: "EUSR",
      description: "Registrations and cards",
      url: props.eusrUrl,
      tone: "charcoal",
      imageUrl: props.eusrImageUrl,
      icon: (
        <Certificate24Regular style={{ width: 40, height: 40 }} aria-hidden />
      ),
    },
    {
      id: "customer-docs",
      title: "Customer Documents",
      description: "Client files by company",
      url: props.customerDocsUrl,
      tone: "green",
      imageUrl: props.customerDocsImageUrl,
      icon: <Folder24Regular style={{ width: 40, height: 40 }} aria-hidden />,
    },
    {
      id: "admin-docs",
      title: "Admin Documents",
      description: "Internal and restricted",
      url: props.adminDocsUrl,
      tone: "charcoal",
      imageUrl: props.adminDocsImageUrl,
      icon: (
        <LockClosed24Regular style={{ width: 40, height: 40 }} aria-hidden />
      ),
    },
  ];

  const onKeyActivate = (
    event: React.KeyboardEvent<HTMLDivElement>,
    url: string,
    title: string
  ): void => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      openUrl(url, title);
    }
  };

  return (
    <section
      className={styles.section}
      aria-label="Quick links"
      style={
        {
          ["--pave-green" as string]: paveTheme.green,
          ["--pave-charcoal" as string]: paveTheme.charcoal,
          ["--pave-charcoal-dark" as string]: paveTheme.charcoalDark,
          ["--pave-font-body" as string]: paveTheme.fontBody,
        } as React.CSSProperties
      }
    >
      <div className={styles.grid}>
        {tiles.map((tile) => {
          const hasImage = Boolean((tile.imageUrl || "").trim());
          return (
            <div
              key={tile.id}
              role="link"
              tabIndex={0}
              aria-label={`${tile.title}: ${tile.description}`}
              className={`${styles.tile} ${
                tile.tone === "green" ? styles.tileGreen : styles.tileCharcoal
              }`}
              style={
                hasImage
                  ? {
                      backgroundImage: `linear-gradient(180deg, rgba(0,0,0,0.15), rgba(0,0,0,0.45)), url(${tile.imageUrl})`,
                      backgroundSize: "cover",
                      backgroundPosition: "center",
                    }
                  : undefined
              }
              onClick={() => openUrl(tile.url, tile.title)}
              onKeyDown={(event) => onKeyActivate(event, tile.url, tile.title)}
            >
              <span className={styles.icon} aria-hidden="true">
                {tile.icon}
              </span>
              <span className={styles.copy}>
                <span className={styles.title}>{tile.title}</span>
                <span className={styles.description}>{tile.description}</span>
              </span>
            </div>
          );
        })}
      </div>
    </section>
  );
};

export default QuickTiles;

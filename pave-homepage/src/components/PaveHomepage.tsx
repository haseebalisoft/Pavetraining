import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { Footer } from "./Footer/Footer";
import { Hero } from "./Hero/Hero";
import { PopularDocs } from "./PopularDocs/PopularDocs";
import { QuickTiles } from "./QuickTiles/QuickTiles";
import { TrainingTracker } from "./TrainingTracker/TrainingTracker";
import { UpcomingEvents } from "./UpcomingEvents/UpcomingEvents";
import { paveTheme } from "../theme/paveTheme";
import styles from "./PaveHomepage.module.scss";

export interface IPaveHomepageProps {
  context: WebPartContext;
  heroHeadline: string;
  heroSubtitle: string;
  greetingNameOverride: string;
  nporsRegisterUrl: string;
  heroImageUrl: string;
  workforceUrl: string;
  eusrUrl: string;
  customerDocsUrl: string;
  adminDocsUrl: string;
  documentsUrl: string;
  eventsAddUrl: string;
  workforceImageUrl: string;
  eusrImageUrl: string;
  customerDocsImageUrl: string;
  adminDocsImageUrl: string;
}

export const PaveHomepage: React.FC<IPaveHomepageProps> = (props) => {
  const documentsSeeAll =
    (props.documentsUrl || "").trim() ||
    (props.customerDocsUrl || "").trim();

  return (
    <div
      className={styles.page}
      style={
        {
          ["--pave-font-body" as string]: paveTheme.fontBody,
          ["--pave-green" as string]: paveTheme.green,
        } as React.CSSProperties
      }
    >
      <div className={styles.container}>
        <Hero
          context={props.context}
          nporsRegisterUrl={props.nporsRegisterUrl}
          heroImageUrl={props.heroImageUrl}
          heroHeadline={props.heroHeadline}
          heroSubtitle={props.heroSubtitle}
          greetingNameOverride={props.greetingNameOverride}
        />

        <div className={styles.stackGap} aria-hidden="true" />

        <QuickTiles
          workforceUrl={props.workforceUrl}
          eusrUrl={props.eusrUrl}
          customerDocsUrl={props.customerDocsUrl}
          adminDocsUrl={props.adminDocsUrl}
          workforceImageUrl={props.workforceImageUrl}
          eusrImageUrl={props.eusrImageUrl}
          customerDocsImageUrl={props.customerDocsImageUrl}
          adminDocsImageUrl={props.adminDocsImageUrl}
        />

        <div className={styles.stackGap} aria-hidden="true" />

        <div className={styles.split}>
          <div className={styles.splitMain}>
            <PopularDocs
              context={props.context}
              documentsUrl={documentsSeeAll}
            />
          </div>
          <div className={styles.splitSide}>
            <UpcomingEvents
              context={props.context}
              eventsAddUrl={props.eventsAddUrl}
            />
          </div>
        </div>

        <div className={styles.stackGap} aria-hidden="true" />

        <TrainingTracker context={props.context} />

        <Footer />
      </div>
    </div>
  );
};

export default PaveHomepage;

import * as React from "react";

import { paveTheme } from "../../theme/paveTheme";
import styles from "./Footer.module.scss";

export const Footer: React.FC = () => {
  return (
    <footer
      className={styles.footer}
      aria-label="PAVE Training footer"
      style={
        {
          ["--pave-green" as string]: paveTheme.green,
          ["--pave-charcoal" as string]: paveTheme.charcoal,
          ["--pave-font-body" as string]: paveTheme.fontBody,
        } as React.CSSProperties
      }
    >
      <div className={styles.monogram} aria-hidden="true">
        <span>P</span>
      </div>
      <p className={styles.tagline}>Paving the way in industry</p>
    </footer>
  );
};

export default Footer;

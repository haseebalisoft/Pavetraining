import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { getGraph } from "../../services/SPContext";
import { paveTheme } from "../../theme/paveTheme";
import styles from "./Hero.module.scss";

export interface IHeroProps {
  context: WebPartContext;
  nporsRegisterUrl: string;
  heroImageUrl: string;
  heroHeadline: string;
  heroSubtitle: string;
  greetingNameOverride: string;
}

interface IGraphMe {
  givenName?: string;
  displayName?: string;
}

function resolveFirstName(
  me: IGraphMe | null,
  fallbackDisplayName: string,
  override: string
): string {
  if (override && override.trim()) {
    return override.trim();
  }
  if (me && me.givenName && me.givenName.trim()) {
    return me.givenName.trim();
  }
  if (me && me.displayName && me.displayName.trim()) {
    return me.displayName.trim().split(/\s+/)[0];
  }
  if (fallbackDisplayName) {
    return fallbackDisplayName.trim().split(/\s+/)[0] || "there";
  }
  return "there";
}

export const Hero: React.FC<IHeroProps> = (props) => {
  const {
    context,
    nporsRegisterUrl,
    heroImageUrl,
    heroHeadline,
    heroSubtitle,
    greetingNameOverride,
  } = props;
  const [firstName, setFirstName] = React.useState<string>(
    greetingNameOverride || "there"
  );

  React.useEffect(() => {
    let cancelled = false;

    if (greetingNameOverride && greetingNameOverride.trim()) {
      setFirstName(greetingNameOverride.trim());
      return;
    }

    (async () => {
      const pageName =
        context.pageContext.user.displayName ||
        context.pageContext.user.loginName ||
        "";

      try {
        const graph = getGraph(context);
        const me = (await graph.me.select(
          "givenName",
          "displayName"
        )()) as IGraphMe;
        if (!cancelled) {
          setFirstName(resolveFirstName(me, pageName, ""));
        }
      } catch (error) {
        console.warn(
          "[Hero] Graph /me unavailable — using page context name",
          error
        );
        if (!cancelled) {
          setFirstName(resolveFirstName(null, pageName, ""));
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [context, greetingNameOverride]);

  const onOpenRegister = React.useCallback((): void => {
    const url = (nporsRegisterUrl || "").trim();
    if (!url) {
      console.warn(
        "[Hero] nporsRegisterUrl is not configured in the property pane"
      );
      return;
    }
    window.location.assign(url);
  }, [nporsRegisterUrl]);

  const onKeyOpen = (event: React.KeyboardEvent<HTMLButtonElement>): void => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      onOpenRegister();
    }
  };

  const imageSrc =
    (heroImageUrl || "").trim() ||
    `${context.pageContext.web.serverRelativeUrl}/SiteAssets/pave-hero-placeholder.jpg`;

  const headline =
    (heroHeadline || "").trim() || "NPORS renewals are due this month";
  const subtitle =
    (heroSubtitle || "").trim() ||
    "Check the register to see which operators need re-certification";

  return (
    <section
      className={styles.hero}
      aria-label="Homepage welcome"
      style={
        {
          ["--pave-green" as string]: paveTheme.green,
          ["--pave-green-dark" as string]: paveTheme.greenDark,
          ["--pave-green-tint" as string]: paveTheme.greenTint,
          ["--pave-charcoal" as string]: paveTheme.charcoal,
          ["--pave-charcoal-dark" as string]: paveTheme.charcoalDark,
          ["--pave-font-heading" as string]: paveTheme.fontHeading,
          ["--pave-font-body" as string]: paveTheme.fontBody,
        } as React.CSSProperties
      }
    >
      <div className={styles.inner}>
        <div className={styles.copy}>
          <p className={styles.greeting}>Welcome back, {firstName} 👋</p>
          <h1 className={styles.headline}>{headline}</h1>
          <p className={styles.subtitle}>{subtitle}</p>
          <button
            type="button"
            className={styles.cta}
            aria-label="Open NPORS register"
            onClick={onOpenRegister}
            onKeyDown={onKeyOpen}
          >
            Open NPORS register
          </button>
        </div>

        <div className={styles.media}>
          <img
            className={styles.image}
            src={imageSrc}
            alt="PAVE Training operations"
          />
        </div>
      </div>
    </section>
  );
};

export default Hero;

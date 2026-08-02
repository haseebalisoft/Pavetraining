"use client";

import Link from "next/link";
import {
  useEffect,
  useMemo,
  useState,
  type CSSProperties,
  type KeyboardEvent,
} from "react";

import type { CustomerOfferRecord } from "@/types/models";

import styles from "./customerDashboard.module.css";

const AUTOPLAY_MS = 7_000;

/**
 * Three static brand heroes (same design + dimensions), then live offers.
 * 1) Brand / paver  2) Accreditations  3) Discount only
 */
const STATIC_SLIDES: CustomerOfferRecord[] = [
  {
    id: "static-brand-hero",
    title: "Paving the way in industry",
    category: "PAVE Training",
    description:
      "Track workforce training, renewals, and bookings for your company — all in one place.",
    startDate: null,
    endDate: null,
    status: "Active",
    image: "/images/customer-hero-paver.png",
    ctaLabel: "Open training matrix",
    ctaLink: "/customer",
  },
  {
    id: "static-accreditations",
    title: "Industry accredited training",
    category: "Accreditations",
    description:
      "NOCN · NPORS · EUSR Approved Trainer · Street Works Qualifications Register — trusted credentials for your workforce.",
    startDate: null,
    endDate: null,
    status: "Active",
    image: "/images/customer-hero-accreditations.png",
    ctaLabel: "View training records",
    ctaLink: "/customer/training-records",
  },
  {
    id: "static-discount",
    title: "15% off NPORS refresher courses",
    category: "Limited time",
    description:
      "Book any NPORS refresher for your team before 30 September and save 15% — valid for all active PAVE Training customers.",
    startDate: null,
    endDate: null,
    status: "Active",
    image: "/images/customer-hero-discount.png",
    ctaLabel: "View offer",
    ctaLink: "/customer/offers",
  },
];

function usePrefersReducedMotion(): boolean {
  const [reduced, setReduced] = useState(false);

  useEffect(() => {
    const query = window.matchMedia("(prefers-reduced-motion: reduce)");
    const update = () => setReduced(query.matches);
    update();
    query.addEventListener("change", update);
    return () => query.removeEventListener("change", update);
  }, []);

  return reduced;
}

function backgroundStyle(image: string | null): CSSProperties {
  const overlay =
    "linear-gradient(100deg, rgba(51, 51, 51, 0.92) 4%, rgba(51, 51, 51, 0.72) 48%, rgba(51, 51, 51, 0.42) 100%)";
  return {
    backgroundImage: image ? `${overlay}, url(${JSON.stringify(image)})` : overlay,
  };
}

export function CustomerOfferSlider({
  offers,
}: {
  offers: CustomerOfferRecord[];
}) {
  const slides = useMemo(
    () => [...STATIC_SLIDES, ...offers],
    [offers],
  );
  const [activeIndex, setActiveIndex] = useState(0);
  const [paused, setPaused] = useState(false);
  const reducedMotion = usePrefersReducedMotion();

  useEffect(() => {
    if (slides.length < 2 || paused || reducedMotion) return;
    const timer = window.setInterval(() => {
      setActiveIndex((current) => (current + 1) % slides.length);
    }, AUTOPLAY_MS);
    return () => window.clearInterval(timer);
  }, [slides.length, paused, reducedMotion]);

  const safeActiveIndex = activeIndex % slides.length;
  const slide = slides[safeActiveIndex] ?? slides[0];
  const goTo = (index: number) => {
    setActiveIndex((index + slides.length) % slides.length);
  };
  const onKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if (event.key === "ArrowLeft") {
      event.preventDefault();
      goTo(safeActiveIndex - 1);
    } else if (event.key === "ArrowRight") {
      event.preventDefault();
      goTo(safeActiveIndex + 1);
    }
  };

  return (
    <section
      className={styles.offerSlider}
      style={backgroundStyle(slide.image)}
      aria-label="Offers and announcements"
      aria-roledescription="carousel"
      tabIndex={0}
      onKeyDown={onKeyDown}
      onMouseEnter={() => setPaused(true)}
      onMouseLeave={() => setPaused(false)}
      onFocusCapture={() => setPaused(true)}
      onBlurCapture={() => setPaused(false)}
    >
      <div key={slide.id} className={styles.offerContent} aria-live="polite">
        <span className={styles.offerTag}>
          {slide.category?.trim() || "Limited time"}
        </span>
        <h2 className={styles.offerTitle}>{slide.title}</h2>
        {slide.description ? (
          <p className={styles.offerDescription}>{slide.description}</p>
        ) : null}
        <Link
          className={styles.offerCta}
          href={slide.ctaLink?.trim() || "/customer/offers"}
        >
          {slide.ctaLabel?.trim() || "View offer"}
          <span aria-hidden="true"> →</span>
        </Link>
      </div>

      {slides.length > 1 ? (
        <>
          <button
            type="button"
            className={`${styles.offerArrow} ${styles.offerArrowPrevious}`}
            aria-label="Previous slide"
            onClick={() => goTo(safeActiveIndex - 1)}
          >
            <span aria-hidden="true">‹</span>
          </button>
          <button
            type="button"
            className={`${styles.offerArrow} ${styles.offerArrowNext}`}
            aria-label="Next slide"
            onClick={() => goTo(safeActiveIndex + 1)}
          >
            <span aria-hidden="true">›</span>
          </button>
          <div className={styles.offerDots} aria-label="Choose slide">
            {slides.map((item, index) => (
              <button
                key={item.id}
                type="button"
                className={
                  index === safeActiveIndex
                    ? styles.offerDotActive
                    : styles.offerDot
                }
                aria-label={`Show slide ${index + 1}: ${item.title}`}
                aria-current={index === safeActiveIndex ? "true" : undefined}
                onClick={() => goTo(index)}
              />
            ))}
          </div>
        </>
      ) : null}
    </section>
  );
}

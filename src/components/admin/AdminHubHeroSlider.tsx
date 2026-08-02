"use client";

import Link from "next/link";
import { useEffect, useState } from "react";

import styles from "./admin.module.css";

export type HubHeroSlide = {
  id: string;
  eyebrow: string;
  title: string;
  subtitle: string;
  ctaLabel: string;
  ctaHref: string;
  metricLabel?: string;
  metricValue?: string | number;
};

interface AdminHubHeroSliderProps {
  slides: HubHeroSlide[];
}

export function AdminHubHeroSlider({ slides }: AdminHubHeroSliderProps) {
  const [index, setIndex] = useState(0);
  const [paused, setPaused] = useState(false);
  const total = slides.length;
  const active = slides[index] ?? slides[0];

  useEffect(() => {
    if (total <= 1 || paused) return;
    const timer = window.setInterval(() => {
      setIndex((current) => (current + 1) % total);
    }, 6500);
    return () => window.clearInterval(timer);
  }, [paused, total]);

  if (!active) return null;

  function goTo(next: number) {
    setIndex((next + total) % total);
  }

  return (
    <div
      className={styles.hubSlider}
      onMouseEnter={() => setPaused(true)}
      onMouseLeave={() => setPaused(false)}
      onFocusCapture={() => setPaused(true)}
      onBlurCapture={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
          setPaused(false);
        }
      }}
    >
      <div className={styles.hubSliderStage} aria-live="polite">
        <div key={active.id} className={styles.hubSliderSlide}>
          <div className={styles.hubHeroCopy}>
            <p className={styles.hubHeroEyebrow}>{active.eyebrow}</p>
            <h1 className={styles.hubHeroTitle}>{active.title}</h1>
            <p className={styles.hubHeroSubtitle}>{active.subtitle}</p>
            <div className={styles.hubSliderActions}>
              <Link className={styles.hubSliderCta} href={active.ctaHref}>
                {active.ctaLabel}
              </Link>
              {active.metricLabel != null && active.metricValue != null ? (
                <div className={styles.hubSliderMetric}>
                  <span className={styles.hubSliderMetricValue}>
                    {active.metricValue}
                  </span>
                  <span className={styles.hubSliderMetricLabel}>
                    {active.metricLabel}
                  </span>
                </div>
              ) : null}
            </div>
          </div>
        </div>
      </div>

      <div className={styles.hubSliderControls}>
        <button
          type="button"
          className={styles.hubSliderNavBtn}
          aria-label="Previous view"
          onClick={() => goTo(index - 1)}
        >
          ‹
        </button>
        <div className={styles.hubSliderDots} role="tablist" aria-label="Hero views">
          {slides.map((slide, slideIndex) => (
            <button
              key={slide.id}
              type="button"
              role="tab"
              aria-selected={slideIndex === index}
              aria-label={`Show ${slide.eyebrow}`}
              className={`${styles.hubSliderDot} ${
                slideIndex === index ? styles.hubSliderDotActive : ""
              }`}
              onClick={() => goTo(slideIndex)}
            />
          ))}
        </div>
        <button
          type="button"
          className={styles.hubSliderNavBtn}
          aria-label="Next view"
          onClick={() => goTo(index + 1)}
        >
          ›
        </button>
      </div>
    </div>
  );
}

import { spawnSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import path from "path";
import withSerwistInit from "@serwist/next";
import type { NextConfig } from "next";

const revision =
  spawnSync("git", ["rev-parse", "HEAD"], { encoding: "utf-8" }).stdout?.trim() ||
  randomUUID();

const withSerwist = withSerwistInit({
  swSrc: "src/app/sw.ts",
  swDest: "public/sw.js",
  disable: process.env.NODE_ENV === "development",
  additionalPrecacheEntries: [{ url: "/~offline", revision }],
});

const nextConfig: NextConfig = {
  turbopack: {
    root: path.join(__dirname),
  },
  experimental: {
    // Reuse recently-visited routes from the client-side router cache so
    // navigating back to a page you were just on is instant instead of
    // re-fetching from SharePoint. `dynamic` defaults to 0 (no reuse) in
    // Next 16, which is why every navigation currently refetches.
    // Values are in seconds; kept modest so admins editing data don't see
    // stale rows for long (server-side SharePoint cache TTL is ~45s).
    staleTimes: {
      // Keep short — admin delete/save must not reappear from router cache.
      dynamic: 0,
      static: 180,
    },
  },
};

export default withSerwist(nextConfig);

/**
 * Resolver hook for the LIVE bulk-upload end-to-end driver
 * (scripts/e2e-workforce-matrix-sync.mjs).
 *
 * Extends scripts/_alias-hook.mjs with the two things needed to run the app's
 * server modules outside Next.js:
 *   - "@/x"          -> <repo>/src/x  (tsconfig paths, same as the unit tests)
 *   - "next/cache"   -> next/cache.js (Next ships these as explicit .js files;
 *                       the bundler adds the extension, plain Node does not)
 *
 * Run with:
 *   node --env-file=.env.local --conditions=react-server \
 *        --import ./scripts/_register-e2e-hook.mjs scripts/e2e-workforce-matrix-sync.mjs
 *
 * `--conditions=react-server` makes the `server-only` package resolve to its
 * empty stub instead of the module that throws.
 */
import { existsSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { resolve as resolvePath } from "node:path";

const SRC = resolvePath(process.cwd(), "src");
const CANDIDATE_SUFFIXES = [
  "",
  ".ts",
  ".tsx",
  ".mjs",
  ".js",
  "/index.ts",
  "/index.tsx",
];

/** Extensionless Next.js subpath entries that plain Node cannot resolve. */
const NEXT_SUBPATHS = new Set([
  "next/headers",
  "next/navigation",
  "next/server",
]);

/**
 * next/cache is STUBBED rather than resolved: unstable_cache requires a
 * request-scoped incremental cache that only exists inside a Next server.
 */
const STUBS = new Map([["next/cache", "./_e2e-next-cache-stub.mjs"]]);

export async function resolve(specifier, context, nextResolve) {
  // Must come before the node_modules fallback below, which would otherwise
  // bypass the package's exports map and load the module that throws.
  if (specifier === "server-only" || specifier === "client-only") {
    const stub = resolvePath(
      process.cwd(),
      "node_modules",
      specifier,
      "empty.js",
    );
    if (existsSync(stub)) {
      return nextResolve(pathToFileURL(stub).href, context);
    }
  }

  const stub = STUBS.get(specifier);
  if (stub) {
    return nextResolve(new URL(stub, import.meta.url).href, context);
  }

  if (specifier.startsWith("@/")) {
    const base = resolvePath(SRC, specifier.slice(2));
    for (const suffix of CANDIDATE_SUFFIXES) {
      const candidate = base + suffix;
      if (existsSync(candidate)) {
        return nextResolve(pathToFileURL(candidate).href, context);
      }
    }
  }

  if (NEXT_SUBPATHS.has(specifier)) {
    const candidate = resolvePath(
      process.cwd(),
      "node_modules",
      `${specifier}.js`,
    );
    if (existsSync(candidate)) {
      return nextResolve(pathToFileURL(candidate).href, context);
    }
  }

  // CommonJS deps whose subpath imports rely on bundler directory resolution
  // (e.g. @microsoft/microsoft-graph-client/authProviders/.../index.js).
  if (!specifier.startsWith(".") && !specifier.startsWith("/")) {
    const asDir = resolvePath(process.cwd(), "node_modules", specifier);
    for (const suffix of ["/index.js", ".js"]) {
      const candidate = asDir + suffix;
      if (existsSync(candidate)) {
        return nextResolve(pathToFileURL(candidate).href, context);
      }
    }
  }

  return nextResolve(specifier, context);
}

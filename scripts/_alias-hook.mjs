/**
 * Minimal ESM resolver hook for the standalone unit tests.
 *
 * Node's native TypeScript type-stripping (>= 22) can load our .ts modules
 * directly, but it does not read tsconfig `paths`, so the project's "@/..."
 * alias fails to resolve. This hook maps "@/x" to "<repo>/src/x" and appends a
 * source extension (or /index) when the specifier is extensionless, matching
 * the bundler resolution used in the app.
 *
 * Register it with:  node --import ./scripts/_register-alias.mjs --test <file>
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

export async function resolve(specifier, context, nextResolve) {
  if (specifier.startsWith("@/")) {
    const base = resolvePath(SRC, specifier.slice(2));
    for (const suffix of CANDIDATE_SUFFIXES) {
      const candidate = base + suffix;
      if (existsSync(candidate)) {
        return nextResolve(pathToFileURL(candidate).href, context);
      }
    }
  }
  return nextResolve(specifier, context);
}

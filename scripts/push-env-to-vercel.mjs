#!/usr/bin/env node
/**
 * Push every key from `.env.local` to the currently-linked Vercel project,
 * across development / preview / production. Skips VERCEL_OIDC_TOKEN (auto).
 *
 * Uses the local vercel CLI (node_modules/.bin/vercel) so no auth token
 * ever leaves the CLI's keystore.
 *
 * Usage: node scripts/push-env-to-vercel.mjs [--file .env.local] [--dry-run]
 */
import { readFileSync } from "node:fs";
import { spawn } from "node:child_process";
import { resolve } from "node:path";

const args = process.argv.slice(2);
const dryRun = args.includes("--dry-run");
const fileIdx = args.indexOf("--file");
const envFile = fileIdx >= 0 ? args[fileIdx + 1] : ".env.local";

const VERCEL_BIN = resolve("node_modules/.bin/vercel");
const ENVIRONMENTS = ["development", "preview", "production"];
const SKIP = new Set(["VERCEL_OIDC_TOKEN"]);

/** Parse `.env.local` files that may have leading-space Vercel-CLI formatting. */
function parseEnvFile(path) {
  const text = readFileSync(path, "utf8");
  const out = new Map();
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.replace(/^\s+/, "");
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq <= 0) continue;
    const key = line.slice(0, eq).trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) continue;
    let value = line.slice(eq + 1);
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    out.set(key, value);
  }
  return out;
}

function runVercel(argv, stdinValue) {
  return new Promise((resolvePromise) => {
    const child = spawn(VERCEL_BIN, argv, { stdio: ["pipe", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (d) => (stdout += d.toString()));
    child.stderr.on("data", (d) => (stderr += d.toString()));
    child.on("close", (code) => resolvePromise({ code, stdout, stderr }));
    child.stdin.end(stdinValue);
  });
}

async function pushOne(key, value, environment) {
  const result = await runVercel(["env", "add", key, environment], value + "\n");
  return result;
}

async function main() {
  const entries = parseEnvFile(envFile);
  const targets = [...entries.entries()].filter(([k]) => !SKIP.has(k));
  console.log(
    `Parsed ${entries.size} vars from ${envFile} (${targets.length} to push, skipping ${SKIP.size}).`,
  );
  if (!targets.length) {
    console.log("Nothing to push.");
    return;
  }
  if (dryRun) {
    for (const [k] of targets) {
      console.log(`DRY: would push ${k} -> ${ENVIRONMENTS.join(", ")}`);
    }
    return;
  }

  let ok = 0;
  let skipped = 0;
  let failed = 0;
  for (const [key, value] of targets) {
    const results = await Promise.all(
      ENVIRONMENTS.map((env) => pushOne(key, value, env).then((r) => [env, r])),
    );
    const status = results.map(([env, r]) => {
      const combined = (r.stdout + r.stderr).toLowerCase();
      if (r.code === 0) {
        ok += 1;
        return `${env}:ok`;
      }
      if (combined.includes("already exists")) {
        skipped += 1;
        return `${env}:exists`;
      }
      failed += 1;
      return `${env}:FAIL(${r.code})`;
    });
    console.log(`${key.padEnd(52)}  ${status.join("  ")}`);
    const failureLog = results.find(
      ([, r]) => r.code !== 0 && !(r.stdout + r.stderr).toLowerCase().includes("already exists"),
    );
    if (failureLog) {
      const [env, r] = failureLog;
      console.log(`  [${env}] stderr: ${r.stderr.trim().slice(0, 300)}`);
    }
  }
  console.log(`\nDone. ok=${ok} exists=${skipped} failed=${failed}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

/**
 * Alias "server-only" to an empty stub for local Node scripts.
 * Usage: node --import ./scripts/register-server-only-stub.mjs ...
 */
import { register } from "node:module";
import { pathToFileURL } from "node:url";
import { resolve } from "node:path";

register(
  pathToFileURL(resolve("./scripts/stubs/server-only-loader.mjs")).href,
  pathToFileURL("./"),
);

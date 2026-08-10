/**
 * Registers scripts/_e2e-hook.mjs for the live bulk-upload E2E driver.
 * See _e2e-hook.mjs for the full invocation.
 */
import { register } from "node:module";

register("./_e2e-hook.mjs", import.meta.url);

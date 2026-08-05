/**
 * Registers the "@/" resolver hook (scripts/_alias-hook.mjs) for the Node test
 * runner. Used via:
 *
 *   node --import ./scripts/_register-alias.mjs --test scripts/test-company-match.mjs
 */
import { register } from "node:module";

register("./_alias-hook.mjs", import.meta.url);

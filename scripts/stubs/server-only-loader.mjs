import { resolve as pathResolve } from "node:path";
import { pathToFileURL } from "node:url";

export async function resolve(specifier, context, nextResolve) {
  if (specifier === "server-only") {
    return {
      shortCircuit: true,
      url: pathToFileURL(pathResolve("./scripts/stubs/server-only.mjs")).href,
    };
  }
  if (specifier.startsWith("@/")) {
    const target = pathResolve("./src", specifier.slice(2));
    const candidates = [
      target,
      `${target}.ts`,
      `${target}.tsx`,
      `${target}.js`,
      `${target}/index.ts`,
    ];
    for (const file of candidates) {
      try {
        const { existsSync } = await import("node:fs");
        if (!existsSync(file)) continue;
        return {
          shortCircuit: true,
          url: pathToFileURL(file).href,
          format: "module",
        };
      } catch {
        // try next
      }
    }
  }
  return nextResolve(specifier, context);
}

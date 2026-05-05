#!/usr/bin/env node
/**
 * Watches resolved *.bas folder and refreshes JSON snapshot (debounced).
 */
import { mkdirSync, watch, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

import { resolveProjectName, resolveStoreOutPath, resolveVbaRoot } from "../dynamicResolve.js";
import { buildIrStoreFromSamplesDir } from "./buildIrStore.js";

function usage(): void {
  console.error(
    [
      "Usage: watch-sample-bas [--samples <dir>] [--out <file>] [--project <name>]",
      "",
      "Defaults: same dynamic resolution as build-sample-store (env + cwd).",
    ].join("\n")
  );
  process.exitCode = 1;
  process.exit();
}

function parseArgs(argv: string[]): {
  samplesDir?: string;
  outPath?: string;
  project?: string;
} {
  const out: { samplesDir?: string; outPath?: string; project?: string } = {};

  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--samples" || a === "--vba") {
      const v = argv[++i];
      if (!v) usage();
      out.samplesDir = v;
      continue;
    }
    if (a === "--out") {
      const v = argv[++i];
      if (!v) usage();
      out.outPath = v;
      continue;
    }
    if (a === "--project") {
      const v = argv[++i];
      if (!v) usage();
      out.project = v;
      continue;
    }
    if (a === "--help" || a === "-h") usage();
  }

  return out;
}

function rebuild(vbaRoot: string, outPath: string, project: string): void {
  const store = buildIrStoreFromSamplesDir(vbaRoot, project);
  mkdirSync(dirname(outPath), { recursive: true });
  writeFileSync(outPath, JSON.stringify(store, null, 2) + "\n", "utf8");
  console.error(
    `[watch-sample-bas] ${vbaRoot} → ${outPath} (${store.procedures.length} procedures, ${store.semantics.length} semantics)`
  );
}

function main(): void {
  const cwd = process.cwd();
  const flags = parseArgs(process.argv.slice(2));

  const vbaRoot = resolveVbaRoot(cwd, flags.samplesDir);
  const project = flags.project ?? resolveProjectName(cwd);
  const outPath = flags.outPath ? resolve(flags.outPath) : resolveStoreOutPath(cwd);

  rebuild(vbaRoot, outPath, project);

  let t: ReturnType<typeof setTimeout> | undefined;
  const schedule = (): void => {
    if (t) clearTimeout(t);
    t = setTimeout(() => {
      try {
        rebuild(vbaRoot, outPath, project);
      } catch (e) {
        console.error("[watch-sample-bas]", e);
      }
    }, 200);
  };

  watch(vbaRoot, { persistent: true }, (_evt, filename) => {
    if (!filename) return;
    if (!filename.toLowerCase().endsWith(".bas")) return;
    schedule();
  });

  console.error(`[watch-sample-bas] watching ${vbaRoot} → ${outPath}`);
}

main();

#!/usr/bin/env node
/**
 * One-shot: scan resolved .bas folder → optional JSON snapshot (paths from env + cwd).
 */
import { mkdirSync, readdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

import { resolveProjectName, resolveStoreOutPath, resolveVbaRoot } from "../dynamicResolve.js";
import { buildIrStoreFromSamplesDir } from "./buildIrStore.js";

function usage(): void {
  console.error(
    [
      "Usage: build-sample-store [--samples <dir>] [--out <file>] [--project <name>]",
      "",
      "Defaults (all dynamic):",
      "  --samples  VQL_VBA_DIR or auto-discover (./samples, ./vba, …)",
      "  --out      VQL_STORE_OUT or ./data/sample-store.json",
      "  --project  VQL_PROJECT or package.json name or cwd folder name",
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

function main(): void {
  const cwd = process.cwd();
  const flags = parseArgs(process.argv.slice(2));

  const vbaRoot = resolveVbaRoot(cwd, flags.samplesDir);
  const project = flags.project ?? resolveProjectName(cwd);
  const outPath = flags.outPath ? resolve(flags.outPath) : resolveStoreOutPath(cwd);

  const basFiles = readdirSync(vbaRoot).filter((f: string) => f.toLowerCase().endsWith(".bas"));
  const store = buildIrStoreFromSamplesDir(vbaRoot, project);

  mkdirSync(dirname(outPath), { recursive: true });
  writeFileSync(outPath, JSON.stringify(store, null, 2) + "\n", "utf8");
  console.error(
    `Wrote ${outPath} from ${vbaRoot} (${basFiles.length} .bas, ${store.procedures.length} procedures, ${store.semantics.length} semantics, ${store.issues.length} issues)`
  );
}

main();

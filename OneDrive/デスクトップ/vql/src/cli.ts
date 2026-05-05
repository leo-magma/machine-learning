#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { basename, extname, resolve } from "node:path";

import { cliOutputExpectsDirectory, resolveCliOutputDestination } from "./cliOutputPath.js";
import {
  type CliOutputFormat,
  inferFormatFromOutputPath,
  serializeQueryResults,
} from "./formatQueryOutput.js";
import { resolveProjectName, resolveVbaRoot } from "./dynamicResolve.js";
import { buildIrStoreFromSamplesDir } from "./extract/buildIrStore.js";
import { parseQuery } from "./parser.js";
import { executeQuery } from "./runQuery.js";
import type { IrStore } from "./store.js";

function usage(): void {
  console.error(
    [
      `Usage:`,
      `  vql [options] '<query>' | path.vq`,
      ``,
      `  Scans *.bas dynamically (each run). Resolution order:`,
      `    --vba / --samples <dir>  →  VQL_VBA_DIR  →  ./samples  →  ./vba  →  ./src/vba  →  .`,
      `  Project name: --project  →  VQL_PROJECT  →  package.json "name"  →  cwd folder`,
      ``,
      `  Options:`,
      `    --store, -s <ir.json>   use a frozen JSON bundle instead of scanning .bas`,
      `    --file, -f <path>       query text file`,
      `    --project, -p <name>    override project id in IR`,
      `    --verbose, -v           print resolved roots to stderr`,
      `    --output, -o <path>     write results to file (suppresses JSON on stdout; prints path on stderr)`,
      `                            directory paths (e.g. report or report/) create the folder if needed`,
      `                            and save each run as …/<name>-<timestamp>-<id>.<ext> (default .html;`,
      `                            <name> is the .vq/.vql basename when used, else vql)`,
      `                            -o report.html also writes under ./report/ (same unique naming, not ./report.html)`,
      `    --format <fmt>          jsonl | pretty | json-array | csv | html (default: jsonl)`,
      `                            with -o, .html/.csv/.json/.jsonl suffix selects format unless overridden`,
      `    --pretty                same as --format pretty`,
      ``,
      `  Example:`,
      `    vql 'FIND procedure LIMIT 5'`,
      `    vql queries/example.vq`,
      `    vql --pretty queries/example.vq`,
      `    vql -o report.html queries/example.vq`,
      `    vql -o out.csv --format csv 'FIND procedure LIMIT 100'`,
    ].join("\n")
  );
  process.exitCode = 1;
  process.exit();
}

function querySourceStemForOutput(cwd: string, queryFile: string | undefined, queryParts: string[]): string | undefined {
  const fromFlag = queryFile?.trim();
  if (fromFlag) {
    return basename(fromFlag).replace(/\.(vq|vql)$/i, "");
  }
  if (queryParts.length === 1) {
    const only = queryParts[0]!.trim();
    if (/\.(vq|vql)$/i.test(only) && existsSync(resolve(cwd, only))) {
      const b = basename(only);
      return b.slice(0, b.length - extname(b).length);
    }
  }
  return undefined;
}

function readQueryTextFromArgs(argvPositional: string[], fileFlag?: string): string {
  if (fileFlag) {
    return readFileSync(resolve(fileFlag), "utf8").replace(/^\uFEFF/, "");
  }
  if (argvPositional.length === 1) {
    const only = argvPositional[0]!;
    if (/\.(vq|vql)$/i.test(only) && existsSync(resolve(only))) {
      return readFileSync(resolve(only), "utf8").replace(/^\uFEFF/, "");
    }
  }
  return argvPositional.join(" ").trim();
}

function loadStore(opts: {
  cwd: string;
  storePath?: string;
  vbaDir?: string;
  project: string;
  verbose?: boolean;
}): IrStore {
  if (opts.vbaDir && opts.storePath) {
    throw new Error("Use either --vba <dir> or --store <json>, not both.");
  }
  if (opts.storePath) {
    const raw = readFileSync(resolve(opts.storePath), "utf8");
    return JSON.parse(raw) as IrStore;
  }

  const root = resolveVbaRoot(opts.cwd, opts.vbaDir);
  if (opts.verbose) {
    console.error(`[vql] vbaRoot=${root} project=${opts.project}`);
  }
  return buildIrStoreFromSamplesDir(root, opts.project);
}

function main(): void {
  const cwd = process.cwd();
  const argv = process.argv.slice(2);
  let storePath: string | undefined;
  let vbaDir: string | undefined;
  let queryFile: string | undefined;
  let projectFromFlag: string | undefined;
  let verbose = false;
  let outputFormat: CliOutputFormat = "jsonl";
  let formatExplicit = false;
  let outputPath: string | undefined;
  const queryParts: string[] = [];

  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--store" || a === "-s") {
      storePath = argv[++i];
      continue;
    }
    if (a === "--vba" || a === "--samples") {
      vbaDir = argv[++i];
      continue;
    }
    if (a === "--project" || a === "-p") {
      const pn = argv[++i];
      if (!pn) usage();
      projectFromFlag = pn;
      continue;
    }
    if (a === "--file" || a === "-f") {
      queryFile = argv[++i];
      continue;
    }
    if (a === "--verbose" || a === "-v") {
      verbose = true;
      continue;
    }
    if (a === "--output" || a === "-o") {
      outputPath = argv[++i];
      continue;
    }
    if (a === "--pretty") {
      outputFormat = "pretty";
      formatExplicit = true;
      continue;
    }
    if (a === "--format") {
      const f = argv[++i];
      formatExplicit = true;
      const allowed: CliOutputFormat[] = ["jsonl", "pretty", "json-array", "csv", "html"];
      if (allowed.includes(f as CliOutputFormat)) outputFormat = f as CliOutputFormat;
      else {
        console.error(`Unknown --format ${f}; use ${allowed.join(", ")}.`);
        process.exitCode = 1;
        process.exit();
      }
      continue;
    }
    if (a === "--help" || a === "-h") usage();
    queryParts.push(a);
  }

  const queryText = readQueryTextFromArgs(queryParts, queryFile).trim();
  if (!queryText) usage();

  const project = projectFromFlag ?? resolveProjectName(cwd);

  let store!: IrStore;
  try {
    store = loadStore({ cwd, storePath, vbaDir, project, verbose });
  } catch (e) {
    console.error(String(e instanceof Error ? e.message : e));
    process.exitCode = 1;
    process.exit();
  }

  const q = parseQuery(queryText);
  const rows = executeQuery(store, q);

  let displayFormat: CliOutputFormat = outputFormat;
  if (outputPath) {
    if (formatExplicit) {
      displayFormat = outputFormat;
    } else {
      const inferred = inferFormatFromOutputPath(outputPath);
      if (inferred !== undefined) displayFormat = inferred;
      else if (cliOutputExpectsDirectory(cwd, outputPath)) displayFormat = "html";
      else displayFormat = outputFormat;
    }
  }

  let payload: string;
  try {
    payload = serializeQueryResults(rows, displayFormat, {
      queryHint: displayFormat === "html" ? queryText.slice(0, 1200) : undefined,
      proceduresForGraph:
        displayFormat === "html"
          ? (store.procedures as Record<string, unknown>[])
          : undefined,
    });
  } catch (e) {
    console.error(String(e instanceof Error ? e.message : e));
    process.exitCode = 1;
    process.exit();
  }

  if (outputPath) {
    const dest = resolveCliOutputDestination(cwd, outputPath, displayFormat, {
      queryStem: querySourceStemForOutput(cwd, queryFile, queryParts),
    });
    writeFileSync(dest, payload, "utf8");
    console.error(`[vql] wrote ${rows.length} row(s) -> ${dest}`);
  } else {
    if (!payload.endsWith("\n")) payload += "\n";
    process.stdout.write(payload);
  }
}

main();

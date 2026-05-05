import { randomBytes } from "node:crypto";
import { existsSync, mkdirSync, statSync } from "node:fs";
import { basename, dirname, extname, join, normalize, resolve } from "node:path";

import { inferFormatFromOutputPath, type CliOutputFormat } from "./formatQueryOutput.js";

export function extensionForCliFormat(format: CliOutputFormat): string {
  switch (format) {
    case "html":
      return "html";
    case "csv":
      return "csv";
    case "json-array":
      return "json";
    case "jsonl":
      return "jsonl";
    case "pretty":
      return "txt";
  }
}

/** `-o report.html` / `report.htm` → actual path `…/<parent>/report/vql-<stamp>-….html` */
function isReportHtmlLegacyOutput(trimmed: string): boolean {
  const base = basename(normalize(trimmed.replace(/[/\\]+$/, "")));
  return /^report\.html?$/i.test(base);
}

/** Strip `.vq`/`.vql` and unsafe characters for use in generated report filenames. */
export function sanitizeQueryStemForFilename(stem: string): string {
  const noExt = stem.replace(/\.(vq|vql)$/i, "");
  const cleaned = noExt
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^\.+|\.+$/g, "")
    .replace(/^_|_$/g, "");
  return cleaned.slice(0, 80);
}

function uniqueFileName(displayFormat: CliOutputFormat, queryStem?: string): string {
  const d = new Date();
  const p = (n: number, w = 2) => String(n).padStart(w, "0");
  const stamp = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}-${p(d.getMinutes())}-${p(d.getSeconds())}_${p(
    d.getMilliseconds(),
    3
  )}`;
  const ext = extensionForCliFormat(displayFormat);
  const safeStem = queryStem ? sanitizeQueryStemForFilename(queryStem) : "";
  const prefix = safeStem ? `${safeStem}-` : "vql-";
  return `${prefix}${stamp}-${randomBytes(3).toString("hex")}.${ext}`;
}

export type ResolveCliOutputOptions = {
  /** Basename stem of the `.vq`/`.vql` file when the query was loaded from disk. */
  queryStem?: string;
};

/** True when `-o` targets a folder (unique filenames per run), not a single file path. */
export function cliOutputExpectsDirectory(cwd: string, outputPath: string): boolean {
  const trimmed = outputPath.trim();
  if (!trimmed) return false;
  if (/[/\\]$/.test(trimmed)) return true;
  const resolved = resolve(cwd, trimmed);
  if (existsSync(resolved)) {
    return statSync(resolved).isDirectory();
  }
  return directoryIntent(trimmed);
}

/** New path without trailing sep: intent is a directory container (not a concrete file path). */
function directoryIntent(trimmed: string): boolean {
  const base = basename(normalize(trimmed.replace(/[/\\]+$/, "")));
  if (inferFormatFromOutputPath(base) !== undefined) return false;
  if (extname(base) !== "") return false;
  return true;
}

/**
 * Resolves where to write CLI output. Directory targets get a unique file name per run
 * (`[<query-stem>-]vql-<date>_<time>_<ms>-<rand>.<ext>` or `vql-…` when no query file).
 */
export function resolveCliOutputDestination(
  cwd: string,
  outputPath: string,
  displayFormat: CliOutputFormat,
  options?: ResolveCliOutputOptions
): string {
  const trimmed = outputPath.trim();
  const resolved = resolve(cwd, trimmed);
  const stem = options?.queryStem;

  if (isReportHtmlLegacyOutput(trimmed)) {
    const dir = join(dirname(resolved), "report");
    mkdirSync(dir, { recursive: true });
    return join(dir, uniqueFileName(displayFormat, stem));
  }

  if (existsSync(resolved)) {
    if (statSync(resolved).isDirectory()) {
      return join(resolved, uniqueFileName(displayFormat, stem));
    }
    mkdirSync(dirname(resolved), { recursive: true });
    return resolved;
  }

  if (/[/\\]$/.test(trimmed)) {
    mkdirSync(resolved, { recursive: true });
    return join(resolved, uniqueFileName(displayFormat, stem));
  }

  if (directoryIntent(trimmed)) {
    mkdirSync(resolved, { recursive: true });
    return join(resolved, uniqueFileName(displayFormat, stem));
  }

  mkdirSync(dirname(resolved), { recursive: true });
  return resolved;
}

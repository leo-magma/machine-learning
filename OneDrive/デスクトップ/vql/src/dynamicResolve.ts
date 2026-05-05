import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { basename, join, resolve } from "node:path";

function isDir(p: string): boolean {
  try {
    return statSync(p).isDirectory();
  } catch {
    return false;
  }
}

/** True if `dir` exists and contains at least one `*.bas` file. */
export function directoryHasBasFiles(dir: string): boolean {
  const abs = resolve(dir);
  if (!isDir(abs)) return false;
  return readdirSync(abs).some((f) => f.toLowerCase().endsWith(".bas"));
}

/**
 * Resolve the folder of exported `.bas` modules.
 * Order: explicit arg → `VQL_VBA_DIR` → first match under cwd among common roots → cwd.
 */
export function resolveVbaRoot(cwd: string, explicit?: string): string {
  if (explicit) {
    const r = resolve(explicit);
    if (!directoryHasBasFiles(r)) {
      throw new Error(`No .bas files in --vba/--samples directory: ${r}`);
    }
    return r;
  }

  const env = process.env.VQL_VBA_DIR?.trim();
  if (env) {
    const r = resolve(env);
    if (!directoryHasBasFiles(r)) {
      throw new Error(`VQL_VBA_DIR points to a folder with no .bas files: ${r}`);
    }
    return r;
  }

  const candidates = [
    join(cwd, "samples"),
    join(cwd, "vba"),
    join(cwd, "src", "vba"),
    cwd,
  ];
  for (const c of candidates) {
    if (directoryHasBasFiles(c)) return resolve(c);
  }

  throw new Error(
    [
      "No .bas files found.",
      "Export modules to ./samples or ./vba, set VQL_VBA_DIR, or pass --vba <dir>.",
      `Tried: ${candidates.map((c) => resolve(c)).join(", ")}`,
    ].join(" ")
  );
}

/** Logical project name for IR rows. `VQL_PROJECT` → package.json name → cwd folder name. */
export function resolveProjectName(cwd: string): string {
  const env = process.env.VQL_PROJECT?.trim();
  if (env) return env;

  try {
    const raw = readFileSync(join(cwd, "package.json"), "utf8");
    const pkg = JSON.parse(raw) as { name?: string };
    if (typeof pkg.name === "string" && pkg.name.trim() !== "") return pkg.name.trim();
  } catch {
    /* optional */
  }

  return basename(resolve(cwd)) || "default";
}

/** Output path for optional JSON snapshot (`build-store` / `watch`). */
export function resolveStoreOutPath(cwd: string): string {
  const env = process.env.VQL_STORE_OUT?.trim();
  if (env) return resolve(env);
  return resolve(join(cwd, "data", "sample-store.json"));
}

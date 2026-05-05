#!/usr/bin/env node
/**
 * CLI shim so `vql` works via npm bin / global link without relying on dist/cli.js shebang alone.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const cli = join(root, "dist", "cli.js");

if (!existsSync(cli)) {
  console.error(
    "vql: dist/cli.js is missing. Run `npm run build` in the repository root, then retry."
  );
  process.exit(1);
}

const child = spawnSync(process.execPath, [cli, ...process.argv.slice(2)], {
  stdio: "inherit",
  env: process.env,
});

process.exit(child.status === null ? 1 : child.status);

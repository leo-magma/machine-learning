# VQL - VBA Query Language

Declarative queries over **Semantic IR**, dependency hints, and procedure/module quality metrics lifted from `.bas` sources.

## Documentation

- **[Design specification](docs/VQL-DESIGN-v1.md)** - syntax, logical model, field namespaces.
- **[Example queries](examples/README.md)** - runnable `.vq` files and catalog.

## Quick start

**One-shot setup** (dependencies, compile, register the `vql` CLI on your PATH):

```bash
npm run setup
```

**Windows PowerShell:** if you see **「スクリプトの実行が無効」** / `PSSecurityException` for `npm.ps1`, use either:

```powershell
npm.cmd run setup
```

Or double-click / run from any shell:

```text
setup.cmd
```

Equivalent to:

```bash
npm install && npm run build && npm link
```

Then run queries from anywhere:

```bash
vql queries/example.vq
vql queries/eod-batch-scope.vq
```

**PowerShell (strict ExecutionPolicy):** global `npm link` installs **`vql.ps1`**, which may be blocked like `npm.ps1`. Use the **`.cmd`** shim:

```powershell
vql.cmd queries\example.vq
vql.cmd queries\eod-batch-scope.vq
```

Or from this repo without relying on global link:

```powershell
node .\bin\vql.mjs queries\example.vq
```

### Without global `npm link`

If you prefer not to link globally:

```bash
npm install
npm run build
npx vql queries/example.vq
```

Or use the npm script:

```bash
npm run vql -- queries/example.vq
```

If `vql` fails with a missing `dist/cli.js` message, run **`npm run build`** (or **`npm run setup`**).

**Readable JSON:** use **`vql.cmd --pretty queries\example.vq`** (or **`--format pretty`**) for indented objects instead of one JSON per line.

**Report files:** write an HTML table plus summary counts (by module / semantic type / issue kind when present), or CSV / JSON. HTML runs use a **`report`** folder with a **unique file per execution** (`report\vql-<timestamp>-….html`). **`-o report.html`** does the same (it does **not** overwrite a single `report.html` in the project root).

```powershell
vql.cmd -o report.html queries\example.vq
vql.cmd -o report queries\example.vq
vql.cmd -o results.csv queries\semantics-dynamic-sql.vq
vql.cmd -o bundle.json queries\modules-by-coupling.vq
```

Formats are inferred from **`.html`**, **`.csv`**, **`.json`** (pretty JSON array), **`.jsonl`**. Override with **`--format html|csv|json-array|jsonl|pretty`**. Rows are not printed to stdout when **`-o`** is used; you get **`[vql] wrote N row(s) -> ...`** on stderr.

HTML reports end with an offline **Visualization** section: SVG bar charts for module / semantic type / issue kind (and procedure counts when the result is not too fragmented).

### `npm` / `node` not recognized (Windows)

If PowerShell says **`npm` is not recognized** (Japanese: コマンドレットとして認識されません), **Node.js is not installed** or your terminal has not picked up **PATH** yet. The Node binary bundled with Cursor does **not** ship `npm`; you need a full Node.js install.

1. Install **[Node.js LTS](https://nodejs.org/)**, **or** in an elevated PowerShell run:
   ```powershell
   winget install OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
   ```
   Approve the UAC prompt if it appears.

2. **Close and reopen** Cursor and the terminal so `PATH` refreshes.

3. Verify:
   ```powershell
   node -v
   npm -v
   ```

4. From this repo root, run **`npm run setup`** again (or **`npm.cmd run setup`** / **`setup.cmd`** in PowerShell when execution policy blocks `npm.ps1`).

### PowerShell: `npm.ps1` blocked (ExecutionPolicy)

Restarting does not fix this; it is a **policy** setting, not PATH.

**Fastest workaround (no policy change):** always invoke the **`.cmd`** shims (same issue affects **`vql`** → use **`vql.cmd`**):

```powershell
npm.cmd run setup
npm.cmd run build
vql.cmd queries\example.vq
npx.cmd vql queries\example.vq
```

Or run **`setup.cmd`** from Explorer or `cmd.exe`.

**Optional fix** (Current user only):

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Then `npm run setup` works like on macOS/Linux.

VBA roots resolve automatically (`samples/`, env `VQL_VBA_DIR`, etc.); see `src/cli.ts` / design doc.

### Starter queries (`queries/`)

| File | Purpose |
|------|---------|
| `example.vq` | Semantic aggregates (`sum`) |
| `procedures-cross-module.vq` | Cross-module outgoing calls |
| `semantics-dynamic-sql.vq` | Inferred dynamic SQL semantics |
| `procedures-risky.vq` | High risk-score procedures |
| `modules-by-coupling.vq` | Modules by coupling degree |
| `eod-batch-scope.vq` | Procedures in `FinBatch_EOD` only |

More patterns live under `examples/queries/` (see `examples/README.md`).

## Regenerate long finance samples

```bash
npm run gen-long-samples
```

Rewrites `samples/FinSQL_ReportingExtract.bas` and `samples/FinAnalytics_CapitalCharges.bas`.

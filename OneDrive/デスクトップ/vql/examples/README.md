# Examples - VQL queries (`examples/queries/`)

From the repository root, either run **`npm run setup`** once (install + build + **`npm link`**), or **`npm install`** + **`npm run build`** only.

Then:

```bash
vql examples/queries/<filename>.vq
```

If you skipped **`npm link`**, use **`npx vql`** or **`node dist/cli.js`** instead of **`vql`**.

`.bas` modules are resolved dynamically, so edits under `samples/` change query results.

The language is a **small filter dialect** built from a subset of SQL-like ideas: only `FIND ... WHERE ... ORDER BY ... LIMIT`. Comments use `#` or `--`; equality is `=` or `==`; `MATCHES "%..."` is the preferred wildcard form (`LIKE` is a synonym).

## Quick catalog

| File | What it surfaces |
|------|------------------|
| **Dependency graph (procedure)** | |
| `dep-cross-module-outgoing.vq` | Procedures ranked by outbound calls to other standard modules |
| `dep-incoming-fan-in.vq` | Hubs / entrypoints with high **incoming** fan-in |
| `dep-unresolved-outgoing.vq` | Edges not resolved (`Application.Run`, etc.) |
| `dep-outgoing-any.vq` | Procedures with at least one outgoing dependency |
| `dep-batch-layer.vq` | Orchestration confined to `FinBatch_EOD` |
| `dep-bridges.vq` | “Bridge” procedures with both outgoing and incoming edges |
| **Dependency & size (module)** | |
| `mod-coupling-degree-desc.vq` | Modules by estimated coupling degree (descending) |
| `mod-avg-cyclomatic-desc.vq` | Modules by average cyclomatic complexity |
| **Quality metrics (procedure)** | |
| `qual-cyclomatic-high.vq` | High cyclomatic complexity |
| `qual-risk-score-high.vq` | Top risk scores |
| `qual-risk-and-complex.vq` | Combined risk and complexity |
| **Semantic IR** | |
| `sem-all-limited.vq` | Semantic nodes (with a row cap) |
| `sem-aggregate-sum.vq` | Aggregate semantics (sum-like) |
| `sem-join-reconcile.vq` | Join / reconciliation-style semantics |
| `sem-lookup-vlookup.vq` | VLOOKUP-style semantics |
| `sem-io-file.vq` | File IO semantics |
| `sem-sql-dynamic.vq` | SQL-like dynamic string / OPENQUERY / MERGE semantics |
| `sem-in-finledger-scope.vq` | Semantics scoped to module `FinLedger_GL` |
| **Issues** | |
| `iss-on-error-resume-next.vq` | `On Error Resume Next` |
| `iss-magic-number.vq` | Magic-number hints |
| `iss-high-severity.vq` | Severity `high` |
| **Names & layers** | |
| `proc-run-prefix.vq` | Procedures whose names look like `Run*` batch entrypoints |
| `proc-fin-modules.vq` | All procedures under `Fin*` modules |
| `proc-finrisk-only.vq` | Procedures in the risk limits module only |
| **Meta / combos** | |
| `combo-hotspot-procedures.vq` | Frequently called + moderately complex |
| `combo-cross-module-and-risk.vq` | Cross-module calls + risk threshold |
| `overview-all-procedures.vq` | All procedures (by module, capped) |
| **Extra (semantic / metrics)** | |
| `sem-origin-finbatch.vq` | Semantic rows originating from the EOD batch module |
| `sem-join-ledger-sheet.vq` | Joins whose right-hand sheet matches subsidiary / GL auxiliary naming |
| `proc-risk-delta-positive.vq` | Rows with positive `riskScore_delta` |
| `proc-changed-recent.vq` | `changedSince` plus complexity (`changedAt` from mtime) |

For field definitions see `docs/VQL-DESIGN-v1.md` and procedure rows as JSON (`dependency.outgoing[]`, `dependency.*Count`, `quality.*`).

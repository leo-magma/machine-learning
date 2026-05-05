# VQL (VBA Query Language) — Design Specification v1

| Item | Value |
|------|--------|
| Document ID | VQL-SPEC-v1 |
| Version | 1.1 |
| Status | Draft (implementation-ready baseline) |
| Audience | IR store implementers, VQL parser/engine developers, VBA asset quality owners |

---

## 0. Overview

**VQL** is a **declarative query language** for querying **Semantic IR** (what the code means) plus attached **structural quality, dependencies, and history**. It assumes VBA has already been lifted into an IR store—not queried as raw text.

This specification defines goals, the logical data model, syntax, field namespaces, result shape, and extension principles for version 1. Physical schemas (JSON Schema, graph DB labels, etc.) are **out of scope**; VQL specifies **logical names** and **meaning of types**.

---

## 1. Goals and Concepts

### 1.1 Goals

1. Treat VBA as **Semantic IR**, not plain text.
2. Provide one surface for:
   - **Find**: enumerate matching locations/units
   - **Analyze (Quality / Dependency)**: filter and rank by complexity, risk, duplication, coupling
   - **Navigate (Where / How / Why)**: every row carries enough identity (`origin.*`, lines) to jump back to source

VQL complements LLM-on-code workflows: it targets **durable, repeatable analysis** of VBA assets under version control.

### 1.2 Example Questions

- List every place this workbook **sums numeric ranges**.
- Find logic that **reconciles Sheet1 against Sheet2**.
- List **high cyclomatic complexity** procedures only.
- Show procedures **recently changed** that also became **higher risk**.

---

## 2. Logical Data Model (Prerequisites)

VQL assumes a populated **IR store** exposing at least these logical entities.

### 2.1 Core Entities

| Entity | Description |
|--------|-------------|
| **Project** | Workbook / VBA project (e.g. `Book1`) |
| **Module** | Standard / class / form module |
| **Procedure** | Sub / Function / Property body |
| **Syntax IR Node** | AST-shaped nodes (not a primary `FIND` target in v1; extension hook) |
| **Semantic IR Node** | Meaning nodes: aggregate, join, lookup, … |
| **Quality Entry** | Structural metrics attached to procedures/modules/etc. |
| **Dependency Edge** | Calls, sheet refs, IO, … (filters/joins in later revisions) |

### 2.2 Representative `semanticType` Values

Align enumeration with your IR catalog.

| `semanticType` | Examples |
|----------------|----------|
| `aggregate` | `sum`, `avg`, `count`, `max`, `min`, … |
| `filter` | conditional row selection |
| `join` | cross-sheet reconcile / two-table logic |
| `lookup` | VLOOKUP-style reference |
| `sort` | ordering |
| `group_by` | grouped aggregation |
| `copy_paste` | copy & paste ranges |
| `format` | formatting |
| `io` | file IO |
| `interop` | other application automation |

Sub-operations such as `operation == "sum"` are defined per `semanticType`.

### 2.3 Example Quality Fields

Dot notation (see §6):

- `quality.complexity.cyclomatic`
- `quality.complexity.maxNestingDepth`
- `quality.dependency.degree`
- `quality.duplication.similarity`
- `quality.risk.riskScore`
- Change helpers: `changedAt`, `changedSince(...)`, deltas such as `quality.risk.riskScore_delta`

---

## 3. Core Syntax

### 3.1 Query Shape

```
FIND <target>
  [IN <scope>]
  [WHERE <condition>]
  [ORDER BY <field> [ASC|DESC]]
  [LIMIT <n>]
```

**Design stance:** VQL is intentionally **smaller than SQL**. There is no `SELECT` list, no joins between IR entities inside the query language, no subqueries, and no grouping—you filter and sort **one row kind** per query. That keeps parsing predictable and errors localized.

| Clause | Meaning |
|--------|---------|
| `<target>` | Row kind returned in v1: `procedure`, `semantic`, `module`, `issue`, … |
| `<scope>` | Restrict scan: `project "Book1"`, `module "Module1"` |
| `<condition>` | Boolean expression over fields and helpers |
| `ORDER BY` | Sort key |
| `LIMIT` | Max rows |

**Lexical rules (v1)**

- String literals use double quotes `"..."`.
- Identifiers use dotted paths for nesting (`source.sheet`).
- **Comments:** line comments start with `#` or `--` and run to end of line (outside strings).
- **Keywords are case-insensitive** (`FIND`, `find`, `Where`, …).
- **Equality:** both `=` and `==` mean equality (same meaning); prefer `==` for visual distinction from SQL if you like.
- **Boolean / null literals:** `true`, `false`, `null` (also case-insensitive).

**Reference corpus:** The shipped `samples/*.bas` macros use **English worksheet tab names** (for example `BatchMonitor`, `GLSubsidiaryLedger`, `SqlParameters`) so worked examples, semantics heuristics, and query catalogs stay easy to read without mixing locales.

**Query file extensions (convention)**

| Extension | Meaning |
|-----------|---------|
| **`.vql`** | Canonical “full name” for saved queries (matches the language name **VQL**). |
| **`.vq`** | Short alias—easy to type and still reads as “VQ…L”. |

Both contain plain UTF‑8 text (optional BOM). Projects should pick one style for consistency; tools accept either.

### 3.2 Example: Sum Aggregates

```
FIND semantic
  WHERE semanticType == "aggregate"
    AND operation == "sum"
```

### 3.3 Example: Touching Sheet1

```
FIND semantic
  WHERE source.sheet == "Sheet1"
```

---

## 4. `FIND` Targets

### 4.1 `FIND semantic`

**Subject**: Semantic IR nodes.

**Typical fields**: `semanticType`, `operation`, `source.sheet`, `source.column`, `source.range`, `target.variable`, `origin.project`, `origin.module`, `origin.procedure`, `origin.lines.start`, `origin.lines.end`.

For `semanticType == "join"`, stores may expose `left.sheet`, `right.sheet`, etc.; if absent, document as unsupported for that dialect.

### 4.2 `FIND procedure`

**Subject**: Procedures. Fields include `name`, `module`, `quality.*`, `changedSince(...)`.

### 4.3 `FIND module`

**Subject**: Modules. Fields include `name`, `quality.dependency.degree`, aggregated complexity.

### 4.4 `FIND issue`

**Subject**: Quality findings. Fields include `kind`, `location.module`, `location.procedure`, `severity`, optional `semanticType`, `similarity` for duplication.

---

## 5. `WHERE` Expressions

### 5.1 Comparison Operators

`=` / `==`, `!=`, `>`, `<`, `>=`, `<=`, `IN (...)`, `NOT IN (...)`.

**Pattern matching (non-SQL name):** use **`MATCHES "pattern"`** with `%` (any substring) and `_` (single character). `LIKE` is accepted as an alias for the same behavior—prefer **`MATCHES`** in new queries so VQL reads less like SQL.

### 5.2 Boolean Operators

`AND`, `OR`, `NOT`, with parentheses. Precedence: `NOT` binds tightest, then `AND`, then `OR`.

### 5.3 Time Helpers

Boolean predicates with ISO date arguments:

- `changedSince("YYYY-MM-DD")`
- `changedBefore("YYYY-MM-DD")`

---

## 6. Field Namespaces

Keep a machine-readable catalog synced to the store. Examples:

**Semantic**: `semanticType`, `operation`, `source.*`, `target.*`, `origin.*`, `left.*`, `right.*` (when provided).

**Quality**: `quality.complexity.*`, `quality.dependency.*`, `quality.duplication.*`, `quality.risk.*`. Array access like `quality.risk.issues[]` is **unspecified** in v1—flatten or add `EXISTS` in implementations.

**History**: `changedAt`, `changedBy`, helpers and `*_delta` fields as available.

---

## 7. Example Queries

```
FIND semantic
  IN project "Book1"
  WHERE semanticType == "aggregate"
    AND operation == "sum"
```

```
FIND semantic
  WHERE semanticType == "join"
    AND left.sheet == "Sheet1"
    AND right.sheet == "Sheet2"
```

```
FIND procedure
  WHERE changedSince("2026-05-01")
    AND quality.complexity.cyclomatic > 12
```

```
FIND procedure
  WHERE quality.risk.riskScore_delta > 0.2
```

```
FIND issue
  WHERE kind == "duplication"
    AND semanticType == "aggregate"
    AND similarity > 0.9
```

---

## 8. Result Rows

Each row must include target kind, stable identity (project/module/procedure/lines as applicable), and attached semantic/quality fragments. Example `FIND semantic` row:

```json
{
  "targetKind": "semantic",
  "semanticType": "aggregate",
  "operation": "sum",
  "source": {
    "sheet": "Sheet1",
    "column": 1,
    "rows": { "from": 1, "to": 10 }
  },
  "targetRef": {
    "kind": "variable",
    "name": "total"
  },
  "origin": {
    "project": "Book1",
    "module": "Module1",
    "procedure": "SumColumn",
    "lines": { "start": 5, "end": 12 }
  },
  "quality": {
    "complexity": { "cyclomatic": 3 },
    "risk": { "riskScore": 0.2 }
  }
}
```

Use `targetKind` / `targetRef` in JSON to avoid clashing with logical `target.*` field paths.

---

## 9. Extensibility

- New semantic kinds: register `semanticType` + fields in the IR catalog; expose to `WHERE` / `ORDER BY`.
- New metrics: add `quality.<ns>.<metric>`.
- Future syntax (v1 **not** required): `SUGGEST`, `EXPLAIN`, `GRAPH` / `DEPENDS ON`, optional `VQL 1` header for versioning.

---

## 10. Deliberately Open in v1

- Normalizing `FIND issue … semanticType` requires store views or denormalized issue rows.
- `LIKE` escaping and case rules.
- Multi-key `ORDER BY`, NULL ordering.
- Coercion rules across number/string/date.
- Optional `FIND sheet`.

---

## 11. Tooling Hooks

| Layer | Role |
|-------|------|
| IR store | Ingest VBA → semantic + quality |
| VQL CLI | Query string → JSON Lines / SARIF |
| Editor | Jump to `origin.lines` from results |
| CI | Gate on `FIND issue WHERE severity >= …` |

Implementation layers: **parser → planner (filters) → serializer**.

---

## 12. Summary

VQL v1 specifies a **declarative language** over **Semantic IR**, **quality**, and **history**, with `FIND` targets for semantics, procedures, modules, and issues—optimized for **repeatable asset governance**, not ad-hoc LLM prompts.

---

## Revision History

| Version | Date | Notes |
|---------|------|-------|
| 1.0 | 2026-05-05 | Initial draft |
| 1.0-en | 2026-05-05 | Full English specification; aligns with reference TypeScript MVP |

---

## 13. Reference implementation (MVP)

This repository ships a minimal **parser + in-memory evaluator** used to validate the syntax and JSON IR shape:

- TypeScript sources under `src/` (`parser`, `evaluator`, `runQuery`, `cli`, `dynamicResolve`).
- **Default workflow:** run from the repo (or any project) root; the CLI **discovers `*.bas` on every run**—no checked-in IR JSON required.

Run (Node.js 18+):

```bash
npm run setup
```

This runs **`npm install`**, **`npm run build`**, and **`npm link`** so the **`vql`** binary is on your PATH. Without linking, run **`npm install`** + **`npm run build`** and invoke **`npx vql`** or **`node dist/cli.js`**.

**Dynamic resolution (no flags):**

1. **VBA folder:** `--vba` / `--samples` if set → else **`VQL_VBA_DIR`** → else first directory under the current working directory that contains at least one **`.bas`**, trying in order: `./samples`, `./vba`, `./src/vba`, `.` (cwd).
2. **Project id:** `--project` if set → else **`VQL_PROJECT`** → else `package.json` **`name`** → else the **cwd folder name** (embedded in IR as `store.project` / `origin.project`).

With **`npm link`** (included in **`npm run setup`**), **`vql`** is on your PATH; otherwise use **`npx vql`** or **`node dist/cli.js`**:

```bash
vql 'FIND procedure LIMIT 5'
vql queries/example.vq
vql -v 'FIND semantic LIMIT 2'    # stderr: resolved vbaRoot + project
```

Optional JSON snapshot (for tools that want a file on disk). Output path: **`VQL_STORE_OUT`** or `./data/sample-store.json`.

```bash
npm run watch-vba       # watches discovered folder → refreshes snapshot
npm run build-store
npm run build-store:py  # Python; same env vars
```

Environment variables:

| Variable | Purpose |
|----------|---------|
| `VQL_VBA_DIR` | Absolute path to folder containing exported `.bas` files |
| `VQL_PROJECT` | Logical project name stored in IR |
| `VQL_STORE_OUT` | Path written by `build-store` / `watch-vba` |

Frozen bundle (CI / regression): **`--store` / `-s`** still loads a JSON file instead of scanning.

```bash
vql --store ./my-bundle.json queries/example.vq
```

**Ingestion** uses cheap text heuristics (not a full VBA compiler).

**Procedure dependency graph (MVP extension):** when ingesting `.bas`, each `FIND procedure` row may include **`dependency.outgoing`** (resolved `Mod.Proc` / `Call` / `Application.Run` edges), plus scalars **`dependency.outgoingCount`**, **`dependency.crossModuleOutgoingCount`**, **`dependency.unresolvedOutgoingCount`**, **`dependency.incomingCount`**, and **`quality.dependency.degree`** for quick filters (e.g. high cross-module coupling, unresolved `Application.Run` targets).

Saved queries use extensions **`.vql`** or **`.vq`** (see §3.1). The CLI prints **JSON Lines** by default (one JSON object per row). Use **`--pretty`** / **`--format pretty`** for indented JSON blocks, or **`-o path.html`** / **`-o path.csv`** for a single file with an HTML report (includes summary breakdowns) or spreadsheet-friendly CSV (**UTF‑8 with BOM**). **`--format json-array`** writes one JSON array.

**Example query pack:** see **`examples/queries/`** (catalog in **`examples/README.md`**) for dependency fan-out/fan-in, module coupling, semantics (`aggregate` / `join` / `lookup` / `io`), issues, and combined hotspot queries—use these to learn which JSON fields appear for each macro shape.

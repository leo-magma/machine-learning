import { getPath } from "./evaluator.js";
import { buildDependencyFigures } from "./reportDependencyGraph.js";

export type CliOutputFormat = "jsonl" | "pretty" | "json-array" | "csv" | "html";

export function inferFormatFromOutputPath(filePath: string): CliOutputFormat | undefined {
  const b = filePath.toLowerCase();
  if (b.endsWith(".html") || b.endsWith(".htm")) return "html";
  if (b.endsWith(".csv")) return "csv";
  if (b.endsWith(".json")) return "json-array";
  if (b.endsWith(".jsonl") || b.endsWith(".ndjson")) return "jsonl";
  return undefined;
}

function flatten(value: unknown, prefix = ""): Record<string, string> {
  const out: Record<string, string> = {};
  if (value === null || value === undefined) {
    if (prefix) out[prefix] = "";
    return out;
  }
  if (typeof value !== "object") {
    out[prefix || "_"] = String(value);
    return out;
  }
  if (Array.isArray(value)) {
    out[prefix || "_"] = JSON.stringify(value);
    return out;
  }
  const keys = Object.keys(value as object);
  if (keys.length === 0) {
    out[prefix || "_"] = "{}";
    return out;
  }
  for (const k of keys) {
    const key = prefix ? `${prefix}.${k}` : k;
    const v = (value as Record<string, unknown>)[k];
    if (v !== null && typeof v === "object" && !Array.isArray(v)) {
      Object.assign(out, flatten(v, key));
    } else if (Array.isArray(v)) {
      out[key] = JSON.stringify(v);
    } else if (v === undefined) {
      out[key] = "";
    } else {
      out[key] = String(v);
    }
  }
  return out;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function csvCell(s: string): string {
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function sortedColumns(rows: Record<string, unknown>[]): string[] {
  const keys = new Set<string>();
  for (const r of rows) {
    for (const k of Object.keys(flatten(r))) keys.add(k);
  }
  return [...keys].sort();
}

function summarizeCounts(rows: Record<string, unknown>[]): {
  byModule: Map<string, number>;
  bySemanticType: Map<string, number>;
  byIssueKind: Map<string, number>;
  byProcedure: Map<string, number>;
} {
  const byModule = new Map<string, number>();
  const bySemanticType = new Map<string, number>();
  const byIssueKind = new Map<string, number>();
  const byProcedure = new Map<string, number>();
  for (const r of rows) {
    const m = getPath(r, "origin.module");
    if (typeof m === "string" && m.length > 0) {
      byModule.set(m, (byModule.get(m) ?? 0) + 1);
    }
    const proc = getPath(r, "origin.procedure");
    if (typeof proc === "string" && proc.length > 0) {
      byProcedure.set(proc, (byProcedure.get(proc) ?? 0) + 1);
    }
    const st = r["semanticType"];
    if (typeof st === "string") bySemanticType.set(st, (bySemanticType.get(st) ?? 0) + 1);
    const kind = r["kind"];
    if (typeof kind === "string") byIssueKind.set(kind, (byIssueKind.get(kind) ?? 0) + 1);
  }
  return { byModule, bySemanticType, byIssueKind, byProcedure };
}

function renderSummaryHtml(title: string, m: Map<string, number>): string {
  if (m.size === 0) return "";
  const rows = [...m.entries()].sort((a, b) => b[1] - a[1]);
  const lis = rows.map(([k, n]) => `<li>${escapeHtml(k)} — ${n}</li>`).join("");
  return `<section class="summary-block"><h3>${escapeHtml(title)}</h3><ul>${lis}</ul></section>`;
}

const BAR_COLORS = ["#3b82f6", "#10b981", "#f59e0b", "#8b5cf6", "#ec4899", "#06b6d4"];

function truncateLabel(s: string, max: number): string {
  if (s.length <= max) return s;
  return s.slice(0, max - 1) + "\u2026";
}

/** Offline SVG horizontal bar charts (no JS / CDN). */
function renderSvgBarChart(title: string, data: Map<string, number>, opts?: { maxBars?: number }): string {
  const maxBars = opts?.maxBars ?? 18;
  const entries = [...data.entries()].sort((a, b) => b[1] - a[1]).slice(0, maxBars);
  if (entries.length === 0) return "";

  const maxVal = Math.max(...entries.map((e) => e[1]), 1);
  const labelW = 240;
  const barMaxW = 460;
  const rowH = 34;
  const sidePad = 12;
  const w = labelW + barMaxW + sidePad * 2 + 44;
  const h = sidePad * 2 + entries.length * rowH;

  let svg = `<figure class="chart-card"><figcaption>${escapeHtml(title)}</figcaption>`;
  svg += `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" role="img" aria-label="${escapeHtml(title)}">`;

  entries.forEach(([label, count], i) => {
    const y = sidePad + i * rowH;
    const barW = Math.max(2, Math.round((count / maxVal) * barMaxW));
    const fill = BAR_COLORS[i % BAR_COLORS.length]!;
    const lab = truncateLabel(label, 38);
    svg += `<text x="${sidePad}" y="${y + 21}" font-size="13" fill="#111827">${escapeHtml(lab)}</text>`;
    svg += `<rect x="${sidePad + labelW}" y="${y + 4}" width="${barW}" height="22" rx="5" fill="${fill}" opacity="0.9"/>`;
    svg += `<text x="${sidePad + labelW + barW + 8}" y="${y + 21}" font-size="13" font-weight="600" fill="#374151">${count}</text>`;
  });

  svg += `</svg></figure>`;
  return svg;
}

function renderVisualizationSection(
  rows: Record<string, unknown>[],
  proceduresForGraph?: Record<string, unknown>[]
): string {
  if (rows.length === 0) return "";

  const { byModule, bySemanticType, byIssueKind, byProcedure } = summarizeCounts(rows);

  const charts: string[] = [];
  if (byModule.size > 0) charts.push(renderSvgBarChart("Rows by module", byModule));
  if (bySemanticType.size > 0) charts.push(renderSvgBarChart("Rows by semantic type", bySemanticType));
  if (byIssueKind.size > 0) charts.push(renderSvgBarChart("Rows by issue kind", byIssueKind));

  const procEntries = [...byProcedure.entries()].sort((a, b) => b[1] - a[1]);
  const topProcShare =
    procEntries.length > 0 ? procEntries[0]![1]! / rows.length : 0;
  if (
    byProcedure.size > 0 &&
    byProcedure.size <= 40 &&
    !(byProcedure.size > 25 && topProcShare < 0.08)
  ) {
    charts.push(renderSvgBarChart("Rows by procedure (top)", byProcedure, { maxBars: 16 }));
  }

  const depFigures =
    proceduresForGraph && proceduresForGraph.length > 0
      ? buildDependencyFigures(rows, proceduresForGraph)
      : "";

  if (charts.length === 0 && !depFigures) return "";

  return `<section class="viz-section" id="visualization">
<h2>Visualization</h2>
<p class="viz-lead">Counts and dependency sketches from this run (SVG only, offline).</p>
${charts.length ? `<div class="chart-grid">${charts.join("\n")}</div>` : ""}
${depFigures ? `<div class="graph-stack">${depFigures}</div>` : ""}
</section>`;
}

export function serializeQueryResults(
  rows: Record<string, unknown>[],
  format: CliOutputFormat,
  opts?: { queryHint?: string; proceduresForGraph?: Record<string, unknown>[] }
): string {
  switch (format) {
    case "jsonl":
      return rows.map((r) => JSON.stringify(r)).join("\n") + (rows.length ? "\n" : "");
    case "pretty":
      return rows.map((r) => JSON.stringify(r, null, 2)).join("\n\n") + (rows.length ? "\n" : "");
    case "json-array":
      return JSON.stringify(rows, null, 2) + "\n";
    case "csv": {
      if (rows.length === 0) return "";
      const cols = sortedColumns(rows);
      const header = cols.map(csvCell).join(",");
      const lines = rows.map((r) => {
        const f = flatten(r);
        return cols.map((c) => csvCell(f[c] ?? "")).join(",");
      });
      return "\uFEFF" + [header, ...lines].join("\r\n") + "\r\n";
    }
    case "html": {
      const hint = opts?.queryHint?.trim() ?? "";
      const { byModule, bySemanticType, byIssueKind } = summarizeCounts(rows);
      const vizSection = renderVisualizationSection(rows, opts?.proceduresForGraph);
      const cols = rows.length ? sortedColumns(rows) : [];
      const thead =
        cols.length > 0 ? `<tr>${cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr>` : "";
      const tbody =
        rows.length > 0
          ? rows
              .map((r) => {
                const f = flatten(r);
                const tds = cols.map((c) => `<td>${escapeHtml(f[c] ?? "")}</td>`).join("");
                return `<tr>${tds}</tr>`;
              })
              .join("\n")
          : "";

      const summaryBlocks = [
        renderSummaryHtml("Rows by origin.module", byModule),
        renderSummaryHtml("Rows by semanticType", bySemanticType),
        renderSummaryHtml("Rows by issue kind", byIssueKind),
      ]
        .filter(Boolean)
        .join("\n");

      return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<title>VQL results (${rows.length} rows)</title>
<style>
body{font-family:system-ui,-apple-system,sans-serif;margin:1.25rem;line-height:1.45;color:#1a1a1a;background:#fafafa;}
h1{font-size:1.25rem;margin-bottom:.25rem;}
pre.query{font-size:.8rem;background:#eee;padding:.6rem .75rem;border-radius:6px;overflow:auto;white-space:pre-wrap;}
.summary-block{margin:.75rem 0;} .summary-block ul{margin:.35rem 0;padding-left:1.25rem;}
.viz-section{margin-top:2.5rem;padding-top:1.25rem;border-top:2px solid #e5e7eb;}
.viz-section h2{font-size:1.1rem;margin:0 0 .35rem;}
.viz-lead{color:#6b7280;font-size:.88rem;margin:0 0 1rem;}
.chart-grid{display:flex;flex-wrap:wrap;gap:1.25rem;align-items:flex-start;}
.chart-card{margin:0;background:#fff;padding:.75rem;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.08);}
.chart-card figcaption{font-size:.82rem;font-weight:600;color:#475569;margin-bottom:.45rem;}
.graph-stack{display:flex;flex-direction:column;gap:2rem;margin-top:1.25rem;}
.graph-card--wide{width:100%;max-width:min(1200px,100%);}
.graph-card--split .graph-split{width:100%;}
.graph-split{display:flex;flex-wrap:wrap;gap:1.25rem 1.75rem;align-items:flex-start;}
.graph-split__viz{flex:1 1 min(720px,calc(100% - 320px));min-width:260px;}
.graph-split__summary{flex:1 1 280px;max-width:420px;min-width:240px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:1rem 1.15rem;box-sizing:border-box;align-self:stretch;}
.graph-split__title{margin:0 0 .55rem;font-size:.98rem;font-weight:700;color:#0f172a;letter-spacing:.02em;}
.graph-split__list{margin:.35rem 0;padding-left:1.15rem;color:#334155;font-size:.875rem;line-height:1.55;}
.graph-split__list li{margin:.4rem 0;}
.graph-split__list strong{color:#1e293b;font-weight:650;}
.graph-split__footer{margin:.65rem 0 0;font-size:.8rem;color:#64748b;line-height:1.45;}
.graph-caption-sub{font-size:.84rem;color:#64748b;margin:0 0 .65rem;line-height:1.45;max-width:100%;}
.graph-card svg{display:block;width:100%;max-width:100%;height:auto;background:linear-gradient(180deg,#fafbff 0%,#f8fafc 100%);border-radius:8px;border:1px solid #e2e8f0;}
.graph-note{font-size:.88rem;color:#475569;margin:.65rem 0 0;line-height:1.45;max-width:100%;}
table{border-collapse:collapse;width:100%;font-size:.82rem;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08);}
th,td{border:1px solid #ddd;padding:.35rem .5rem;text-align:left;vertical-align:top;}
th{background:#f0f4f8;position:sticky;top:0;}
tr:nth-child(even){background:#f9fafb;}
.meta{color:#555;font-size:.9rem;margin-bottom:1rem;}
</style>
</head>
<body>
<h1>VQL results</h1>
<p class="meta"><strong>${rows.length}</strong> row(s)</p>
${hint ? `<pre class="query">${escapeHtml(hint)}</pre>` : ""}
${summaryBlocks}
${cols.length ? `<table>\n<thead>${thead}</thead>\n<tbody>\n${tbody}\n</tbody>\n</table>` : "<p>No rows.</p>"}
${vizSection}
</body>
</html>
`;
    }
    default: {
      const x: never = format;
      throw new Error(`Unknown format: ${String(x)}`);
    }
  }
}

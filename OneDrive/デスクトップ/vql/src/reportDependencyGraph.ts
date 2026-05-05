import { getPath } from "./evaluator.js";

type OutEdge = { targetModule: string; targetProcedure: string; resolved: boolean };

/** Collect modules referenced by query result rows. */
export function collectFocusModules(rows: Record<string, unknown>[]): Set<string> {
  const s = new Set<string>();
  for (const r of rows) {
    const m = getPath(r, "origin.module");
    if (typeof m === "string" && m.length > 0) s.add(m);
    const rm = r["module"];
    if (typeof rm === "string" && rm.length > 0) s.add(rm);
    const lm = getPath(r, "location.module");
    if (typeof lm === "string" && lm.length > 0) s.add(lm);
  }
  return s;
}

/** Expand with one hop neighbors (callers + callees at module granularity). Cap size. */
export function expandModulesForGraph(
  procedures: Record<string, unknown>[],
  focus: Set<string>,
  maxNodes: number
): Set<string> {
  if (focus.size === 0) return new Set();
  const out = new Set(focus);

  for (const p of procedures) {
    const mod = String(p.module ?? "");
    if (!focus.has(mod)) continue;
    const dep = p.dependency as { outgoing?: OutEdge[] } | undefined;
    for (const e of dep?.outgoing ?? []) {
      if (e.resolved && e.targetModule !== mod) out.add(e.targetModule);
    }
  }

  for (const p of procedures) {
    const mod = String(p.module ?? "");
    const dep = p.dependency as { outgoing?: OutEdge[] } | undefined;
    for (const e of dep?.outgoing ?? []) {
      if (e.resolved && focus.has(e.targetModule) && mod !== e.targetModule) out.add(mod);
    }
  }

  if (out.size <= maxNodes) return out;

  const ranked = [...out].map((name) => ({
    name,
    pri: focus.has(name) ? 10000 : 0,
    deg: moduleTotalDegree(procedures, name),
  }));
  ranked.sort((a, b) => b.pri + b.deg - (a.pri + a.deg));
  return new Set(ranked.slice(0, maxNodes).map((x) => x.name));
}

function moduleTotalDegree(procedures: Record<string, unknown>[], module: string): number {
  let d = 0;
  for (const p of procedures) {
    if (String(p.module ?? "") !== module) continue;
    const dep = p.dependency as { outgoing?: OutEdge[] } | undefined;
    d += (dep?.outgoing ?? []).filter((e) => e.resolved && e.targetModule !== module).length;
  }
  for (const p of procedures) {
    const mod = String(p.module ?? "");
    if (mod === module) continue;
    const dep = p.dependency as { outgoing?: OutEdge[] } | undefined;
    for (const e of dep?.outgoing ?? []) {
      if (e.resolved && e.targetModule === module && mod !== module) d++;
    }
  }
  return d;
}

/** Directed aggregate edges between modules (resolved calls only). */
export function weightedModuleEdges(
  procedures: Record<string, unknown>[],
  modules: Set<string>
): { from: string; to: string; w: number }[] {
  const acc = new Map<string, number>();
  for (const p of procedures) {
    const from = String(p.module ?? "");
    if (!modules.has(from)) continue;
    const dep = p.dependency as { outgoing?: OutEdge[] } | undefined;
    for (const e of dep?.outgoing ?? []) {
      if (!e.resolved) continue;
      const to = e.targetModule;
      if (!modules.has(to) || from === to) continue;
      const k = `${from}\x00${to}`;
      acc.set(k, (acc.get(k) ?? 0) + 1);
    }
  }
  return [...acc.entries()]
    .map(([k, w]) => {
      const [from, to] = k.split("\x00");
      return { from: from!, to: to!, w };
    })
    .sort((a, b) => b.w - a.w);
}

function procPairKey(mod: string, proc: string): string {
  return `${mod}\x00${proc}`;
}

export function collectFocusProcedurePairs(rows: Record<string, unknown>[]): Map<string, string> {
  /** key -> display label */
  const m = new Map<string, string>();
  for (const r of rows) {
    const mod = getPath(r, "origin.module");
    const proc = getPath(r, "origin.procedure");
    if (typeof mod !== "string" || typeof proc !== "string") continue;
    if (!mod || !proc) continue;
    const k = procPairKey(mod, proc);
    if (!m.has(k)) m.set(k, `${mod}.${proc}`);
  }
  for (const r of rows) {
    const mo = r["module"];
    const na = r["name"];
    if (typeof mo === "string" && typeof na === "string" && na) {
      const k = procPairKey(mo, na);
      if (!m.has(k)) m.set(k, `${mo}.${na}`);
    }
  }
  return m;
}

/** Edges between procedures that both appear in focusPairs (resolved only). */
export function procedureEdgesWithinFocus(
  procedures: Record<string, unknown>[],
  focusPairs: Map<string, string>
): { from: string; to: string; label: string }[] {
  const edges: { from: string; to: string; label: string }[] = [];
  const keys = new Set(focusPairs.keys());

  for (const p of procedures) {
    const mod = String(p.module ?? "");
    const name = String(p.name ?? "");
    const fk = procPairKey(mod, name);
    if (!keys.has(fk)) continue;

    const dep = p.dependency as { outgoing?: OutEdge[] } | undefined;
    for (const e of dep?.outgoing ?? []) {
      if (!e.resolved) continue;
      const tk = procPairKey(e.targetModule, e.targetProcedure);
      if (!keys.has(tk)) continue;
      const fromLabel = focusPairs.get(fk) ?? fk;
      const toLabel = focusPairs.get(tk) ?? tk;
      edges.push({
        from: fromLabel,
        to: toLabel,
        label: `${e.targetModule}.${e.targetProcedure}`,
      });
    }
  }
  return edges;
}

function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Stable hash for per-edge curve / label offsets. */
function pairIntHash(a: string, b: string): number {
  let h = 5381;
  for (let i = 0; i < a.length; i++) h = (h * 33) ^ a.charCodeAt(i);
  h = (h * 33) ^ 10;
  for (let i = 0; i < b.length; i++) h = (h * 33) ^ b.charCodeAt(i);
  return h | 0;
}

/** Split long module names across two lines (prefer underscore) to avoid truncation. */
function moduleLabelLines(name: string): string[] {
  const maxSingle = 20;
  if (name.length <= maxSingle) return [name];
  const u = name.indexOf("_");
  if (u >= 3 && u < name.length - 3) {
    const head = name.slice(0, u);
    const tail = name.slice(u + 1);
    if (head.length <= 22 && tail.length <= 22) return [head, tail];
    return [
      head.length > 22 ? `${head.slice(0, 19)}…` : head,
      tail.length > 22 ? `${tail.slice(0, 19)}…` : tail,
    ];
  }
  const mid = Math.ceil(name.length / 2);
  return [name.slice(0, mid), name.slice(mid)];
}

/** Place label centroid outside the node, away from graph center (reduces overlap with the disc). */
function outwardLabelAnchor(
  px: number,
  py: number,
  gcx: number,
  gcy: number,
  nr: number,
  gap: number
): { x: number; y: number } {
  const dx = px - gcx;
  const dy = py - gcy;
  const len = Math.hypot(dx, dy) || 1;
  const push = nr + gap;
  return { x: px + (dx / len) * push, y: py + (dy / len) * push };
}

function shortenSegment(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  shave: number
): { x1: number; y1: number; x2: number; y2: number } {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const len = Math.hypot(dx, dy);
  if (len < shave * 2 + 4) return { x1, y1, x2, y2 };
  const ux = dx / len;
  const uy = dy / len;
  return {
    x1: x1 + ux * shave,
    y1: y1 + uy * shave,
    x2: x2 - ux * shave,
    y2: y2 - uy * shave,
  };
}

function circlePositions(
  nodes: string[],
  cx: number,
  cy: number,
  r: number,
  phase = 0
): Map<string, { x: number; y: number }> {
  const pos = new Map<string, { x: number; y: number }>();
  const n = nodes.length;
  if (n === 0) return pos;
  const sorted = [...nodes].sort((a, b) => a.localeCompare(b));
  for (let i = 0; i < n; i++) {
    const ang = (i / n) * Math.PI * 2 - Math.PI / 2 + phase;
    pos.set(sorted[i]!, { x: cx + r * Math.cos(ang), y: cy + r * Math.sin(ang) });
  }
  return pos;
}

/** Inner ring = modules that appear in query rows; outer = 1-hop neighbours only (clearer semantics). */
function moduleTieredPositions(
  nodes: string[],
  focusModules: Set<string>,
  cx: number,
  cy: number,
  rInner: number,
  rOuter: number
): Map<string, { x: number; y: number }> {
  const sorted = [...nodes].sort((a, b) => a.localeCompare(b));
  const inner = sorted.filter((n) => focusModules.has(n));
  const outer = sorted.filter((n) => !focusModules.has(n));
  if (inner.length > 0 && outer.length > 0) {
    const pos = new Map<string, { x: number; y: number }>();
    const phase = inner.length ? Math.PI / inner.length : 0;
    for (const [k, v] of circlePositions(inner, cx, cy, rInner, 0)) pos.set(k, v);
    for (const [k, v] of circlePositions(outer, cx, cy, rOuter, phase)) pos.set(k, v);
    return pos;
  }
  const r = outer.length === 0 ? rOuter * 0.88 : Math.max(rInner * 1.15, rOuter * 0.75);
  return circlePositions(sorted, cx, cy, r, 0);
}

function svgArrowMarker(uniqueId: string, fill = "#1e293b"): { defs: string; href: string } {
  const id = `arrow-${uniqueId}`;
  /* userSpaceOnUse keeps arrowhead size fixed; strokeWidth scaling made thick edges unreadable */
  return {
    defs: `<defs><marker id="${id}" viewBox="0 0 10 7" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto" markerUnits="userSpaceOnUse"><polygon points="0 0, 10 3.5, 0 7" fill="${fill}"/></marker></defs>`,
    href: `url(#${id})`,
  };
}

function moduleGraphLegend(x: number, y: number, hasNeighbours: boolean): string {
  const line = (dy: number, body: string) =>
    `<g transform="translate(${x}, ${y + dy})"><text x="0" y="5" dominant-baseline="middle">${body}</text></g>`;
  let g = `<g class="graph-legend" font-size="12" fill="#334155">`;
  g += `<rect x="${x - 8}" y="${y - 10}" width="300" height="${hasNeighbours ? 124 : 94}" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5"/>`;
  g += `<text x="${x}" y="${y + 6}" font-weight="700" font-size="13" fill="#0f172a">How to read</text>`;
  g += line(22, `<tspan font-weight="600" fill="#b45309">●</tspan><tspan> = module in your result table</tspan>`);
  if (hasNeighbours) {
    g += line(44, `<tspan font-weight="600" fill="#475569">●</tspan><tspan> = neighbour (1-hop, resolved calls only)</tspan>`);
  }
  g += line(
    hasNeighbours ? 66 : 44,
    `<tspan font-weight="600">→</tspan><tspan> caller module → callee module</tspan>`
  );
  g += line(
    hasNeighbours ? 88 : 66,
    `<tspan>Thicker line / bold number = more procedure calls (aggregated)</tspan>`
  );
  g += `</g>`;
  return g;
}

function renderModuleGraphEnglishAside(opts: {
  moduleCount: number;
  focusCount: number;
  neighbourCount: number;
  drawnEdges: { from: string; to: string; w: number }[];
}): string {
  const { moduleCount, focusCount, neighbourCount, drawnEdges } = opts;
  let top: { from: string; to: string; w: number } | undefined;
  for (const e of drawnEdges) {
    if (!top || e.w > top.w) top = e;
  }
  const totalCalls = drawnEdges.reduce((s, e) => s + e.w, 0);

  const bullets: string[] = [
    `Lists <strong>${moduleCount} modules</strong>, expanded one hop from modules referenced in your query (resolved cross-module calls only).`,
  ];
  if (neighbourCount > 0 && focusCount > 0) {
    bullets.push(
      `<strong>${focusCount}</strong> appear in your result rows; <strong>${neighbourCount}</strong> are neighbours kept only as structural context.`
    );
  } else if (focusCount === moduleCount && moduleCount > 0) {
    bullets.push(`Every module drawn is referenced directly by rows in this report.`);
  } else {
    bullets.push(`Extra modules appear when they directly call—or are called by—a module tied to your rows.`);
  }
  bullets.push(
    `<strong>A → B</strong> means procedures in module <strong>A</strong> reach resolved targets in module <strong>B</strong>; the count aggregates IR outgoing edges between that pair.`
  );
  bullets.push(`Stroke thickness and the numeric label both encode volume—compare pairs using the numbers when lines overlap.`);
  if (top && drawnEdges.length > 0) {
    bullets.push(
      `Heaviest pairwise traffic shown: <strong>${escapeXmlText(top.from)} → ${escapeXmlText(top.to)}</strong> (${top.w}).`
    );
  }
  bullets.push(`Intra-module calls and unresolved references are excluded.`);

  const footer =
    drawnEdges.length > 0 && totalCalls > 0
      ? `<p class="graph-split__footer">Sum of counts on drawn arrows: <strong>${totalCalls}</strong> (display capped at 56 arrows).</p>`
      : "";

  const lis = bullets.map((b) => `<li>${b}</li>`).join("");
  return `<aside class="graph-split__summary"><h4 class="graph-split__title">Quick summary</h4><ul class="graph-split__list">${lis}</ul>${footer}</aside>`;
}

function renderProcedureGraphEnglishAside(procCount: number, uniqueEdgeCount: number): string {
  const bullets = [
    `Each box is one of <strong>${procCount} procedures</strong> present together in this query result.`,
    `Arrows show direct <strong>resolved</strong> calls between those procedures; anything outside this result set is hidden.`,
    uniqueEdgeCount === 0
      ? `No direct edges among these procedures in this snapshot.`
      : uniqueEdgeCount === 1
        ? `There is <strong>one</strong> caller→callee link among them.`
        : `There are <strong>${uniqueEdgeCount}</strong> caller→callee links drawn.`,
    `Use the module graph for aggregate coupling across whole modules.`,
  ];
  const lis = bullets.map((b) => `<li>${b}</li>`).join("");
  return `<aside class="graph-split__summary graph-split__summary--proc"><h4 class="graph-split__title">Quick summary</h4><ul class="graph-split__list">${lis}</ul></aside>`;
}

export function renderModuleDependencySvg(
  modules: Set<string>,
  edges: { from: string; to: string; w: number }[],
  queryModules?: Set<string>
): string {
  const nodeList = [...modules];
  if (nodeList.length === 0) return "";

  const focusInGraph = queryModules ? new Set([...queryModules].filter((m) => modules.has(m))) : new Set<string>();
  const hasNeighbours = nodeList.some((n) => !focusInGraph.has(n));

  const W = 980;
  const H = 720;
  const cx = W / 2 + 24;
  const cy = H / 2 + 8;
  const rInner = Math.min(W, H) * 0.22;
  const rOuter = Math.min(W, H) * 0.38;
  const pos =
    focusInGraph.size > 0 && hasNeighbours
      ? moduleTieredPositions(nodeList, focusInGraph, cx, cy, rInner, rOuter)
      : circlePositions(nodeList, cx, cy, Math.min(W, H) * 0.34);
  const nr = focusInGraph.size > 0 && hasNeighbours ? 30 : 34;

  const mk = svgArrowMarker("mod", "#1e293b");
  const drawn = edges.filter((e) => modules.has(e.from) && modules.has(e.to) && e.from !== e.to).slice(0, 56);
  const neighbourCount = nodeList.filter((n) => !focusInGraph.has(n)).length;

  let svg = `<figure class="chart-card graph-card graph-card--wide graph-card--split"><figcaption>Module dependency graph</figcaption>`;
  svg += `<p class="graph-caption-sub">Orange-filled modules appear in your result; grey modules are neighbours only (one hop). Arrows follow resolved calls (counts aggregated per module pair).</p>`;
  svg += `<div class="graph-split">`;
  svg += `<div class="graph-split__viz">`;
  svg += `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img" aria-label="Module dependency graph">`;
  svg += mk.defs;

  const drawnSorted = [...drawn].sort((a, b) => `${a.from}\x00${a.to}`.localeCompare(`${b.from}\x00${b.to}`));

  for (let ei = 0; ei < drawnSorted.length; ei++) {
    const e = drawnSorted[ei]!;
    const p1 = pos.get(e.from);
    const p2 = pos.get(e.to);
    if (!p1 || !p2) continue;
    const seg = shortenSegment(p1.x, p1.y, p2.x, p2.y, nr + 5);
    const mx = (seg.x1 + seg.x2) / 2;
    const my = (seg.y1 + seg.y2) / 2;
    const h = pairIntHash(e.from, e.to);
    const bendMag = 0.062 + (Math.abs(h) % 160) / 2200;
    const bendSide = h % 2 === 0 ? 1 : -1;
    const ox = -(seg.y2 - seg.y1) * bendMag * bendSide;
    const oy = (seg.x2 - seg.x1) * bendMag * bendSide;
    /* slight fan-out so bundles split toward graph periphery */
    const fan = (ei % 5) * 0.012 * bendSide;
    const cxq = mx + ox + (seg.x2 - seg.x1) * fan;
    const cyq = my + oy + (seg.y2 - seg.y1) * fan;
    const sw = Math.min(3.6, 1.25 + Math.min(e.w, 80) * 0.028);
    svg += `<path d="M ${seg.x1.toFixed(1)} ${seg.y1.toFixed(1)} Q ${cxq.toFixed(1)} ${cyq.toFixed(1)} ${seg.x2.toFixed(1)} ${seg.y2.toFixed(1)}" fill="none" stroke="#94a3b8" stroke-opacity="0.9" stroke-linecap="round" stroke-width="${sw.toFixed(
      2
    )}" marker-end="${mk.href}"/>`;

    const chord = Math.hypot(seg.x2 - seg.x1, seg.y2 - seg.y1) || 1;
    const nx = -(seg.y2 - seg.y1) / chord;
    const ny = (seg.x2 - seg.x1) / chord;
    const bump = 16 + (Math.abs(pairIntHash(`${e.from}|lbl`, e.to)) % 6) * 5;
    const side = pairIntHash(e.to, e.from) % 2 === 0 ? 1 : -1;
    const lx = cxq + nx * bump * side;
    const ly = cyq + ny * bump * side;
    const countStr = String(e.w);
    const pillW = Math.max(22, countStr.length * 9 + 10);
    svg += `<g class="edge-weight-label">`;
    svg += `<rect x="${(lx - pillW / 2).toFixed(1)}" y="${(ly - 9).toFixed(1)}" width="${pillW.toFixed(1)}" height="17" rx="5" fill="#ffffff" fill-opacity="0.94" stroke="#e2e8f0" stroke-width="1"/>`;
    svg += `<text x="${lx.toFixed(1)}" y="${(ly + 3.5).toFixed(1)}" font-size="11" font-weight="700" fill="#0f172a" text-anchor="middle">${countStr}</text>`;
    svg += `</g>`;
  }

  for (const name of nodeList) {
    const p = pos.get(name)!;
    const inFocus = focusInGraph.has(name);
    const fill = inFocus ? "#fef3c7" : "#f1f5f9";
    const stroke = inFocus ? "#b45309" : "#475569";
    svg += `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${nr}" fill="${fill}" stroke="${stroke}" stroke-width="2.5"/>`;
  }

  for (const name of nodeList) {
    const p = pos.get(name)!;
    const inFocus = focusInGraph.has(name);
    const txFill = inFocus ? "#78350f" : "#1e293b";
    const lines = moduleLabelLines(name);
    const gap = lines.length > 1 ? 22 : 18;
    const tip = outwardLabelAnchor(p.x, p.y, cx, cy, nr, gap);
    const halo = `paint-order="stroke fill" stroke="#ffffff" stroke-width="4"`;
    if (lines.length === 1) {
      svg += `<text x="${tip.x.toFixed(1)}" y="${(tip.y + 5).toFixed(1)}" font-size="12.5" font-weight="700" text-anchor="middle" fill="${txFill}" ${halo}>${escapeXmlText(
        lines[0]!
      )}</text>`;
    } else {
      svg += `<text x="${tip.x.toFixed(1)}" y="${(tip.y - 4).toFixed(1)}" font-size="11" font-weight="700" text-anchor="middle" fill="${txFill}" ${halo}>${escapeXmlText(
        lines[0]!
      )}</text>`;
      svg += `<text x="${tip.x.toFixed(1)}" y="${(tip.y + 10).toFixed(1)}" font-size="11" font-weight="600" text-anchor="middle" fill="${txFill}" ${halo}>${escapeXmlText(
        lines[1]!
      )}</text>`;
    }
  }

  svg += moduleGraphLegend(Math.min(W - 16, Math.max(24, W - 308)), Math.min(H - 24, Math.max(140, H - 138)), hasNeighbours);

  svg += `</svg>`;
  svg += `</div>`;
  svg += renderModuleGraphEnglishAside({
    moduleCount: nodeList.length,
    focusCount: focusInGraph.size,
    neighbourCount,
    drawnEdges: drawn,
  });
  svg += `</div>`;
  svg += `<p class="graph-note">Built from IR <code>dependency.outgoing</code> (resolved only). Modules outside your table appear when they are direct callers or callees of a module that is in the result.</p>`;
  svg += `</figure>`;
  return svg;
}

function procGraphLegend(x: number, y: number): string {
  return `<g class="graph-legend" font-size="12" fill="#334155">
  <rect x="${x - 8}" y="${y - 10}" width="292" height="72" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5"/>
  <text x="${x}" y="${y + 6}" font-weight="700" font-size="13" fill="#0f172a">How to read</text>
  <text x="${x}" y="${y + 28}" font-weight="600">→</text><text x="${x + 16}" y="${y + 28}">caller procedure  →  callee (resolved)</text>
  <text x="${x}" y="${y + 50}">Only calls between procedures that both appear in this result.</text>
</g>`;
}

/** Procedure-level graph when few procedures in result (labels shortened). */
export function renderProcedureDependencySvg(
  pairLabels: Map<string, string>,
  edges: { from: string; to: string; label: string }[]
): string {
  if (pairLabels.size === 0 || pairLabels.size > 22) return "";

  const nodes = [...new Set(pairLabels.values())];
  const W = 980;
  const H = Math.max(560, 140 + nodes.length * 58);
  const cx = W / 2 + 20;
  const cy = H / 2 + 16;
  const R = Math.min(W * 0.39, H * 0.39);
  const pos = circlePositions(nodes, cx, cy, R);

  const mk = svgArrowMarker("proc", "#5b21b6");
  const edgeSeen = new Set<string>();
  for (const e of edges) {
    edgeSeen.add(`${e.from}\x00${e.to}`);
  }
  const uniqueEdgeCount = edgeSeen.size;
  edgeSeen.clear();

  let svg = `<figure class="chart-card graph-card graph-card--wide graph-card--split"><figcaption>Procedure calls within result rows</figcaption>`;
  svg += `<p class="graph-caption-sub">Each box is one procedure from your table; arrows are direct resolved calls.</p>`;
  svg += `<div class="graph-split">`;
  svg += `<div class="graph-split__viz">`;
  svg += `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img" aria-label="Procedure dependency graph">`;
  svg += mk.defs;

  for (const e of edges) {
    const key = `${e.from}\x00${e.to}`;
    if (edgeSeen.has(key)) continue;
    edgeSeen.add(key);
    const p1 = pos.get(e.from);
    const p2 = pos.get(e.to);
    if (!p1 || !p2) continue;
    const seg = shortenSegment(p1.x, p1.y, p2.x, p2.y, 40);
    svg += `<line x1="${seg.x1.toFixed(1)}" y1="${seg.y1.toFixed(1)}" x2="${seg.x2.toFixed(1)}" y2="${seg.y2.toFixed(1)}" stroke="#7c3aed" stroke-width="2.5" stroke-opacity="0.9" marker-end="${mk.href}"/>`;
  }

  for (const label of nodes) {
    const p = pos.get(label)!;
    const parts = label.split(".");
    const modRaw = parts[0] ?? "";
    const procRaw = parts.slice(1).join(".") || "(procedure)";
    const line1 = modRaw.length > 20 ? modRaw.slice(0, 18) + "\u2026" : modRaw;
    const line2 = procRaw.length > 24 ? procRaw.slice(0, 22) + "\u2026" : procRaw;
    const tw = Math.min(200, 14 + Math.max(line1.length, line2.length) * 7.2);
    const th = 44;
    const x0 = p.x - tw / 2;
    const y0 = p.y - th / 2;
    svg += `<rect x="${x0.toFixed(1)}" y="${y0.toFixed(1)}" width="${tw.toFixed(1)}" height="${th}" rx="10" fill="#faf5ff" stroke="#6d28d9" stroke-width="2.5"/>`;
    svg += `<text x="${p.x.toFixed(1)}" y="${(p.y - 7).toFixed(1)}" font-size="12.5" font-weight="700" text-anchor="middle" fill="#4c1d95">${escapeXmlText(
      line1
    )}</text>`;
    svg += `<text x="${p.x.toFixed(1)}" y="${(p.y + 11).toFixed(1)}" font-size="12" text-anchor="middle" fill="#5b21b6">${escapeXmlText(line2)}</text>`;
  }

  svg += procGraphLegend(22, 26);

  svg += `</svg>`;
  svg += `</div>`;
  svg += renderProcedureGraphEnglishAside(nodes.length, uniqueEdgeCount);
  svg += `</div>`;
  svg += `<p class="graph-note">If this graph is empty, your procedures may not call each other directly—use the module graph above for cross-module structure.</p>`;
  svg += `</figure>`;
  return svg;
}

export function buildDependencyFigures(
  queryRows: Record<string, unknown>[],
  procedures: Record<string, unknown>[] | undefined
): string {
  if (!procedures?.length) return "";

  const focus = collectFocusModules(queryRows);
  const pairs = collectFocusProcedurePairs(queryRows);

  const parts: string[] = [];

  if (focus.size > 0) {
    const expandedMods = expandModulesForGraph(procedures, focus, 30);
    const wEdges = weightedModuleEdges(procedures, expandedMods);
    if (wEdges.length > 0 && expandedMods.size >= 2) {
      parts.push(renderModuleDependencySvg(expandedMods, wEdges, focus));
    } else if (expandedMods.size >= 2) {
      parts.push(
        `<figure class="chart-card graph-card"><figcaption>Module neighbourhood</figcaption><p class="graph-note">Expanded modules (${escapeXmlText([...expandedMods].sort().join(", "))}) but no resolved cross-module calls among them in the IR.</p></figure>`
      );
    }
  }

  if (pairs.size >= 2 && pairs.size <= 22) {
    const pedges = procedureEdgesWithinFocus(procedures, pairs);
    if (pedges.length > 0) {
      parts.push(renderProcedureDependencySvg(pairs, pedges));
    }
  }

  if (parts.length === 0 && focus.size === 0 && pairs.size === 0) {
    parts.push(
      `<figure class="chart-card graph-card"><figcaption>Dependency graphs</figcaption><p class="graph-note">No <code>origin.module</code> / procedure identity in rows to anchor graphs. Try <code>FIND procedure ... LIMIT 40</code> or queries whose rows include <code>origin.*</code>.</p></figure>`
    );
  }

  return parts.join("\n");
}

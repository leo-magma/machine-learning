import { readFileSync, statSync } from "node:fs";

export type ProcedureScan = {
  name: string;
  kind: "Sub" | "Function";
  startLine: number;
  endLine: number;
  body: string;
};

export type ModuleScan = {
  path: string;
  moduleName: string;
  lines: string[];
  procedures: ProcedureScan[];
  mtimeIso: string;
};

const VB_NAME_RE = /^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/i;
const PROC_START_RE = /^\s*(?:Public|Private)?\s*(Sub|Function)\s+(\w+)\s*(?:\(|$)/i;
const PROC_END_RE = /^\s*End\s+(Sub|Function)\s*$/i;

function stripComments(line: string): string {
  const q = line.indexOf("'");
  if (q === -1) return line;
  return line.slice(0, q);
}

export function readModuleBas(absPath: string): ModuleScan {
  const raw = readFileSync(absPath, "utf8").replace(/^\uFEFF/, "");
  const lines = raw.split(/\r?\n/);
  let moduleName = basenameToModule(absPath);
  for (const ln of lines.slice(0, 30)) {
    const m = ln.match(VB_NAME_RE);
    if (m?.[1]) {
      moduleName = m[1];
      break;
    }
  }

  const st = statSync(absPath);
  const mtimeIso = new Date(st.mtimeMs).toISOString();

  const procedures = extractProcedures(lines);

  return { path: absPath, moduleName, lines, procedures, mtimeIso };
}

function basenameToModule(p: string): string {
  const base = p.replace(/^.*[/\\]/, "").replace(/\.bas$/i, "");
  return base || "UnknownModule";
}

function extractProcedures(lines: string[]): ProcedureScan[] {
  type Frame = { name: string; kind: "Sub" | "Function"; startIdx: number };
  const stack: Frame[] = [];
  const out: ProcedureScan[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? "";
    const sm = line.match(PROC_START_RE);
    if (sm?.[1] && sm[2]) {
      stack.push({
        name: sm[2],
        kind: sm[1].toLowerCase() === "function" ? "Function" : "Sub",
        startIdx: i,
      });
      continue;
    }
    const em = line.match(PROC_END_RE);
    if (em && stack.length > 0) {
      const top = stack.pop()!;
      const endIdx = i;
      const bodyLines = lines.slice(top.startIdx, endIdx + 1);
      const body = bodyLines.join("\n");
      out.push({
        name: top.name,
        kind: top.kind,
        startLine: top.startIdx + 1,
        endLine: endIdx + 1,
        body,
      });
    }
  }

  return out;
}

/** Rough cyclomatic-style complexity: decision / loop heads in VBA-ish code. */
export function estimateCyclomatic(body: string): number {
  const text = body
    .split(/\r?\n/)
    .map((l) => stripComments(l))
    .join("\n");

  const patterns: RegExp[] = [
    /\bIf\b/gi,
    /\bElseIf\b/gi,
    /\bElse\b/gi,
    /\bSelect\s+Case\b/gi,
    /\bCase\b/gi,
    /\bFor\s+Each\b/gi,
    /\bFor\b/gi,
    /\bWhile\b/gi,
    /\bDo\s+(?:While|Until)?\b/gi,
  ];

  let d = 0;
  for (const re of patterns) {
    const m = text.match(re);
    d += m?.length ?? 0;
  }
  return Math.max(1, d);
}

export function listMagicNumberIssues(
  moduleName: string,
  proc: ProcedureScan,
  project: string
): Record<string, unknown>[] {
  const issues: Record<string, unknown>[] = [];
  const bodyLines = proc.body.split(/\r?\n/);
  const skipRe =
    /^\s*(Attribute|Option|Dim|Const|Private\s+Const|Public\s+Const|Enum\b|Declare\b)/i;
  const maxIssues = 40;

  for (let i = 0; i < bodyLines.length; i++) {
    if (issues.length >= maxIssues) break;
    const physicalLine = proc.startLine + i;
    let ln = stripComments(bodyLines[i] ?? "");
    if (skipRe.test(ln)) continue;
    const matches = ln.match(/\b(\d{2,})\b/g);
    if (!matches) continue;
    for (const num of matches) {
      if (issues.length >= maxIssues) break;
      if (num === "10" || num === "16" || num === "32" || num === "64") continue;
      issues.push({
        kind: "magic_number",
        severity: Number(num) >= 500 ? "medium" : "low",
        detail: { literal: Number(num) },
        location: {
          module: moduleName,
          procedure: proc.name,
          lines: { start: physicalLine, end: physicalLine },
        },
        origin: { project, module: moduleName, procedure: proc.name },
      });
    }
  }
  return issues;
}

export function sheetsMentioned(body: string): string[] {
  const found = new Set<string>();
  const text = stripComments(body.replace(/\r/g, "\n"));

  const reWs = /(?:Worksheets|Sheets)\s*\(\s*"([^"]+)"\s*\)/gi;
  let m: RegExpExecArray | null;
  while ((m = reWs.exec(text))) found.add(m[1]);

  const reBare = /\b(Sheet\d+)\b/gi;
  while ((m = reBare.exec(text))) found.add(m[1]);

  return [...found];
}

export function hasOnErrorResumeNext(body: string): boolean {
  return /^\s*On\s+Error\s+Resume\s+Next\s*$/im.test(body);
}

import type { ProcedureScan } from "./scanBas.js";
import { estimateCyclomatic, sheetsMentioned } from "./scanBas.js";

export type SemanticHit = Record<string, unknown>;

/** Removes text inside VBA `"..."` literals (`""` = escaped "). SQL fragments live there and must not trigger worksheet Sum/Avg heuristics. */
function stripVbaStringLiterals(src: string): string {
  let out = "";
  let i = 0;
  while (i < src.length) {
    const ch = src[i]!;
    if (ch !== '"') {
      out += ch;
      i++;
      continue;
    }
    i++;
    while (i < src.length) {
      if (src[i] === '"') {
        if (src[i + 1] === '"') {
          i += 2;
          continue;
        }
        i++;
        break;
      }
      i++;
    }
    out += " ";
  }
  return out;
}

function riskFromSignals(cc: number, onErr: boolean, magicHints: number): number {
  let r = Math.min(1, cc / 45 + (onErr ? 0.25 : 0) + Math.min(0.35, magicHints * 0.02));
  return Math.round(r * 100) / 100;
}

export function inferSemanticsForProcedure(
  project: string,
  moduleName: string,
  proc: ProcedureScan
): SemanticHit[] {
  const body = proc.body;
  const hits: SemanticHit[] = [];
  const sheets = sheetsMentioned(body);
  const cc = estimateCyclomatic(body);
  const onErr = /On\s+Error\s+Resume\s+Next/i.test(body);
  const magicHints = (body.match(/\b\d{2,}\b/g) ?? []).length;

  const qBase = {
    origin: {
      project,
      module: moduleName,
      procedure: proc.name,
      lines: { start: proc.startLine, end: proc.endLine },
    },
    quality: {
      complexity: { cyclomatic: cc },
      risk: { riskScore: riskFromSignals(cc, onErr, magicHints) },
    },
  };

  const codeSansStrings = stripVbaStringLiterals(body);

  const sumSignals =
    /WorksheetFunction\.Sum\b/i.test(body) ||
    /\bApplication\.WorksheetFunction\.Sum\b/i.test(body) ||
    /(?:^|[^A-Za-z0-9_])Sum\s*\(/i.test(codeSansStrings) ||
    (/\w+\s*=\s*\w+\s*\+\s*.+Cells/i.test(codeSansStrings) &&
      /\bTotal\b|\bRunning\b|sum/i.test(codeSansStrings));

  const avgSignals =
    /WorksheetFunction\.Average\b/i.test(body) ||
    /\bApplication\.Average\s*\(/i.test(codeSansStrings) ||
    /\bAverage\s*\(/i.test(codeSansStrings) ||
    /\bAvg\s*\(/i.test(codeSansStrings);

  const financeReconcileHint =
    /NOSTRO|general ledger|subsidiary ledger|ledger|settlement|reconcile|GLEntry|CORE_GL|nostro balance|journal/i.test(
      body
    );
  const joinSignals =
    sheets.length >= 2 &&
    (/\bFor\s+/i.test(body) || /\bFor\s+Each\b/i.test(body)) &&
    (financeReconcileHint ||
      (sheets.some((s) => /sheet1/i.test(s)) && sheets.some((s) => /sheet2/i.test(s))));

  const lookupSignals =
    /\bVLookup\b/i.test(body) ||
    /\bVLOOKUP\b/i.test(body) ||
    /WorksheetFunction\.VLookup\b/i.test(body);

  const ioSignals =
    /\bOpen\s+".*"?\s+For\b/i.test(body) ||
    /\bCreateObject\s*\(\s*"Scripting\.FileSystemObject"\s*\)/i.test(body);

  const sqlSignals =
    (/\bSELECT\b[\s\S]{0,6000}\bFROM\b/i.test(body) &&
      (/ADODB|CommandText|\.Execute\b|OPENQUERY|MERGE\s+INTO|INSERT\s+INTO|UPDATE\s+dbo\.|EXEC\s+sp_/i.test(
        body
      ) ||
        /&\s*_\s*(\r?\n|\r)/.test(body))) ||
    (/staging|OPENQUERY|linked\s*server/i.test(body) && /\bSELECT\b/i.test(body) && /\bFROM\b/i.test(body));

  if (sumSignals) {
    const sheet = pickSheet(sheets, /sheet1|revenue|inv|NOSTRO|journal|subsidiary|balance/i, "Sheet1");
    hits.push({
      semanticType: "aggregate",
      operation: "sum",
      source: { sheet, column: guessColumn(body, 1), rows: { from: 2, to: guessLastRow(body) } },
      targetRef: { kind: "variable", name: guessSumVar(body) },
      ...qBase,
    });
  }

  if (avgSignals && !sumSignals) {
    hits.push({
      semanticType: "aggregate",
      operation: "avg",
      source: { sheet: pickSheet(sheets, /sheet1/i, "Sheet1"), column: guessColumn(body, 5) },
      targetRef: { kind: "variable", name: "avgResult" },
      ...qBase,
    });
  }

  if (joinSignals) {
    const left =
      sheets.find((s) => /NOSTRO|sheet1|GLSubsidiary|Journal|Subsidiary|Nostro/i.test(s)) ??
      sheets[0] ??
      "Sheet1";
    const right =
      sheets.find((s) => s !== left && /CORE|sheet2|Subsidiary|Journal|Ledger|GL/i.test(s)) ??
      sheets.find((s) => s !== left) ??
      sheets[1] ??
      "Sheet2";
    hits.push({
      semanticType: "join",
      operation: "reconcile",
      left: { sheet: left, keyColumn: guessKeyColumn(body, 1) },
      right: { sheet: right, keyColumn: guessKeyColumn(body, 1) },
      ...qBase,
    });
  }

  if (lookupSignals) {
    hits.push({
      semanticType: "lookup",
      operation: "vlookup",
      source: { sheet: sheets[sheets.length - 1] ?? "Sheet2", column: 3 },
      targetRef: { kind: "range", name: "LookupResults" },
      ...qBase,
    });
  }

  if (ioSignals) {
    hits.push({
      semanticType: "io",
      operation: "file",
      source: { path: "inferred" },
      ...qBase,
    });
  }

  if (sqlSignals) {
    const dialect = /\bOPENQUERY\b/i.test(body)
      ? "linked_server_openquery"
      : /MERGE\s+INTO/i.test(body)
        ? "tsql_merge_staging"
        : "dynamic_adodb";
    hits.push({
      semanticType: "sql",
      operation: "dynamic_text",
      source: {
        dialect,
        worksheetParams: sheets[0] ?? "SqlParameters",
      },
      targetRef: { kind: "connection", name: "ADODB_or_LINKEDSRV" },
      ...qBase,
    });
  }

  return hits;
}

function pickSheet(sheets: string[], prefer: RegExp, fallback: string): string {
  const hit = sheets.find((s) => prefer.test(s));
  return hit ?? sheets[0] ?? fallback;
}

function guessColumn(body: string, def: number): number {
  const m = body.match(/Cells\s*\(\s*\w+\s*,\s*(\d+)\s*\)/i);
  if (m?.[1]) return Number(m[1]);
  return def;
}

function guessKeyColumn(body: string, def: number): number {
  const m = body.match(/Cells\s*\(\s*\w+\s*,\s*(\d+)\s*\)/gi);
  if (m && m[0]) {
    const n = m[0].match(/,\s*(\d+)/);
    if (n?.[1]) return Number(n[1]);
  }
  return def;
}

function guessLastRow(body: string): number {
  const m = body.match(/\b(?:lastRow|LastRow)\b\s*=\s*(\d+)/i);
  if (m?.[1]) return Number(m[1]);
  return 200;
}

function guessSumVar(body: string): string {
  const dim = body.match(/\bDim\s+(\w+)\s+As\s+Double/i);
  if (dim?.[1]) return dim[1];
  const z = body.match(/\b(\w+)\s*=\s*0\b/);
  if (z?.[1] && !/^(Then|Next|If)$/i.test(z[1])) return z[1];
  return "runningTotal";
}

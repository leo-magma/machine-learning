/** Best-effort extraction of macro/procedure references from VBA source (not a full parser). */

export type ResolvedCallEdge = {
  targetModule: string;
  targetProcedure: string;
  resolved: boolean;
  raw: string;
};

const SKIP_OBJECT_PREFIX = new Set([
  "Sheet",
  "Sheets",
  "Worksheets",
  "ChartObjects",
  "Workbook",
  "Workbooks",
  "Range",
  "Cells",
  "Rows",
  "Columns",
  "Application",
  "Debug",
  "VBA",
  "Err",
  "Strings",
  "Math",
  "Interaction",
  "WorksheetFunction",
  "ActiveSheet",
  "ActiveWorkbook",
  "ThisWorkbook",
  "Window",
  "Chart",
  "Shape",
  "Connection",
  "Recordset",
  "Command",
  "Field",
  "Fields",
  "Parameter",
  "Parameters",
]);

/** Member access that is almost never a cross-module procedure call */
const SKIP_MEMBER_SUFFIX = new Set([
  "Cells",
  "Rows",
  "Columns",
  "Range",
  "Select",
  "Activate",
  "Value",
  "Value2",
  "Formula",
  "FormulaR1C1",
  "Clear",
  "Copy",
  "Paste",
  "PasteSpecial",
  "End",
  "Offset",
  "Resize",
  "Interior",
  "Font",
  "Name",
  "Address",
  "Row",
  "Column",
  "Count",
  "EntireRow",
  "EntireColumn",
  "AutoFilter",
  "Find",
  "Open",
  "Close",
  "Execute",
  "MoveNext",
  "MoveFirst",
  "AddNew",
  "Update",
  "State",
  "BOF",
  "EOF",
  "AppendChunk",
  "GetChunk",
]);

function stripLineComments(body: string): string {
  return body
    .split(/\r?\n/)
    .map((line) => {
      const q = line.indexOf("'");
      return q === -1 ? line : line.slice(0, q);
    })
    .join("\n");
}

export type RawOutgoingCall = {
  targetModule?: string;
  targetProcedure: string;
  raw: string;
};

export function extractRawOutgoingCalls(body: string): RawOutgoingCall[] {
  const text = stripLineComments(body);
  const seen = new Set<string>();
  const out: RawOutgoingCall[] = [];

  const push = (tm: string | undefined, tp: string, raw: string): void => {
    const key = `${tm ?? ""}::${tp}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ targetModule: tm, targetProcedure: tp, raw });
  };

  for (const m of text.matchAll(/Application\.Run\s+"([^"]+)"/gi)) {
    const inner = m[1] ?? "";
    const tail = inner.includes("!") ? inner.split("!").pop()! : inner;
    const qp = tail.replace(/^'+|'+$/g, "");
    const dot = qp.lastIndexOf(".");
    if (dot > 0) {
      const mod = qp.slice(0, dot).trim();
      const proc = qp.slice(dot + 1).trim();
      push(mod, proc, `Application.Run "${inner}"`);
    } else if (qp.length > 0) {
      push(undefined, qp.trim(), `Application.Run "${inner}"`);
    }
  }

  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    const callStmt = /^\s*Call\s+(\w+)/i.exec(trimmed);
    if (callStmt?.[1]) {
      push(undefined, callStmt[1], trimmed);
    }
  }

  const modProcRe = /\b([A-Za-z][A-Za-z0-9_]*)\.([A-Za-z][A-Za-z0-9_]*)\b/g;
  let mp: RegExpExecArray | null;
  while ((mp = modProcRe.exec(text))) {
    const obj = mp[1];
    const memb = mp[2];
    if (SKIP_OBJECT_PREFIX.has(obj)) continue;
    if (SKIP_MEMBER_SUFFIX.has(memb)) continue;
    if (/^Sheet\d+$/i.test(obj)) continue;
    push(obj, memb, `${obj}.${memb}`);
  }

  return out;
}

export function resolveOutgoingCalls(
  sourceModule: string,
  raw: RawOutgoingCall[],
  moduleToProcedures: Map<string, Set<string>>,
  procedureNameToModules: Map<string, string[]>
): ResolvedCallEdge[] {
  const edges: ResolvedCallEdge[] = [];

  for (const c of raw) {
    let tm = c.targetModule;
    const tp = c.targetProcedure;

    let resolved = false;
    if (tm && moduleToProcedures.get(tm)?.has(tp)) {
      resolved = true;
    } else if (!tm) {
      const local = moduleToProcedures.get(sourceModule);
      if (local?.has(tp)) {
        tm = sourceModule;
        resolved = true;
      } else {
        const mods = procedureNameToModules.get(tp);
        if (mods?.length === 1) {
          tm = mods[0];
          resolved = true;
        }
      }
    }

    edges.push({
      targetModule: tm ?? "?",
      targetProcedure: tp,
      resolved,
      raw: c.raw,
    });
  }

  return edges;
}

export function crossModuleCount(sourceModule: string, edges: ResolvedCallEdge[]): number {
  let n = 0;
  for (const e of edges) {
    if (!e.resolved) continue;
    if (e.targetModule !== sourceModule) n++;
  }
  return n;
}

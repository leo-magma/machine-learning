import type { Expr } from "./ast.js";

export function getPath(obj: unknown, path: string): unknown {
  const parts = path.split(".");
  let cur: unknown = obj;
  for (const p of parts) {
    if (cur === null || cur === undefined) return undefined;
    if (typeof cur !== "object") return undefined;
    cur = (cur as Record<string, unknown>)[p];
  }
  return cur;
}

export function evalWhere(row: unknown, expr: Expr | undefined): boolean {
  if (!expr) return true;
  return toBoolean(evalExpr(row, expr));
}

function toBoolean(v: unknown): boolean {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0 && !Number.isNaN(v);
  if (typeof v === "string") return v.length > 0;
  return v != null;
}

export function evalExpr(row: unknown, expr: Expr): unknown {
  switch (expr.type) {
    case "literal":
      return expr.value;
    case "field":
      return getPath(row, expr.path);
    case "call":
      return evalCall(row, expr.name, expr.date);
    case "unary":
      if (expr.op === "NOT") return !toBoolean(evalExpr(row, expr.expr));
      throw new Error("Unknown unary");
    case "binary": {
      if (expr.op === "AND")
        return toBoolean(evalExpr(row, expr.left)) && toBoolean(evalExpr(row, expr.right));
      if (expr.op === "OR")
        return toBoolean(evalExpr(row, expr.left)) || toBoolean(evalExpr(row, expr.right));
      const l = evalExpr(row, expr.left);
      const r = evalExpr(row, expr.right);
      return cmp(expr.op, l, r);
    }
    case "in": {
      const v = evalExpr(row, expr.expr);
      return expr.values.some((x) => looseEq(v, x));
    }
    case "notIn": {
      const v = evalExpr(row, expr.expr);
      return !expr.values.some((x) => looseEq(v, x));
    }
    case "like": {
      const v = evalExpr(row, expr.left);
      return likeMatch(v, expr.pattern);
    }
    default:
      throw new Error("Unhandled expr");
  }
}

function evalCall(row: unknown, name: "changedSince" | "changedBefore", date: string): boolean {
  const changedAt = getPath(row, "changedAt");
  if (typeof changedAt !== "string") return false;
  const rowDay = changedAt.slice(0, 10);
  if (name === "changedSince") return rowDay >= date.slice(0, 10);
  return rowDay < date.slice(0, 10);
}

function cmp(op: string, l: unknown, r: unknown): boolean {
  const ln = coerceNum(l);
  const rn = coerceNum(r);
  if (ln !== null && rn !== null) {
    switch (op) {
      case "==":
        return ln === rn;
      case "!=":
        return ln !== rn;
      case ">":
        return ln > rn;
      case ">=":
        return ln >= rn;
      case "<":
        return ln < rn;
      case "<=":
        return ln <= rn;
      default:
        throw new Error(`Unknown numeric op ${op}`);
    }
  }
  const ls = l === null || l === undefined ? "" : String(l);
  const rs = r === null || r === undefined ? "" : String(r);
  switch (op) {
    case "==":
      return ls === rs;
    case "!=":
      return ls !== rs;
    case ">":
      return ls > rs;
    case ">=":
      return ls >= rs;
    case "<":
      return ls < rs;
    case "<=":
      return ls <= rs;
    default:
      throw new Error(`Unknown string op ${op}`);
  }
}

function coerceNum(v: unknown): number | null {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v))) return Number(v);
  return null;
}

function looseEq(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  const an = coerceNum(a);
  const bn = coerceNum(b);
  if (an !== null && bn !== null) return an === bn;
  return String(a ?? "") === String(b ?? "");
}

function likeMatch(value: unknown, pattern: string): boolean {
  const s = value === null || value === undefined ? "" : String(value);
  let out = "^";
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i]!;
    if (c === "%") out += ".*";
    else if (c === "_") out += ".";
    else if (/[*+?^${}()|[\]\\]/.test(c)) out += "\\" + c;
    else out += c;
  }
  out += "$";
  return new RegExp(out, "i").test(s);
}

/** Compare two rows by dotted path for ORDER BY */
export function compareByField(rowA: unknown, rowB: unknown, field: string, dir: "ASC" | "DESC"): number {
  const va = getPath(rowA, field);
  const vb = getPath(rowB, field);
  const na = coerceNum(va);
  const nb = coerceNum(vb);
  let c = 0;
  if (na !== null && nb !== null) c = na - nb;
  else c = String(va ?? "").localeCompare(String(vb ?? ""));
  return dir === "DESC" ? -c : c;
}

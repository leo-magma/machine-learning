import type { Query } from "./ast.js";
import { compareByField, evalWhere } from "./evaluator.js";
import type { IrStore } from "./store.js";

function attachDefaults(store: IrStore, rows: Record<string, unknown>[]): Record<string, unknown>[] {
  return rows.map((r) => ({
    ...r,
    origin: {
      ...(typeof r.origin === "object" && r.origin !== null ? (r.origin as object) : {}),
      project:
        getOrigin(r, "project") ??
        (typeof r.project === "string" ? r.project : undefined) ??
        store.project,
    },
  }));
}

function getOrigin(row: Record<string, unknown>, key: string): unknown {
  const o = row.origin;
  if (typeof o !== "object" || o === null) return undefined;
  return (o as Record<string, unknown>)[key];
}

export function executeQuery(store: IrStore, q: Query): Record<string, unknown>[] {
  let rows: Record<string, unknown>[] = [];
  switch (q.target) {
    case "semantic":
      rows = attachSemantics(store);
      break;
    case "procedure":
      rows = attachDefaults(store, store.procedures as Record<string, unknown>[]);
      break;
    case "module":
      rows = attachDefaults(store, store.modules as Record<string, unknown>[]);
      break;
    case "issue":
      rows = attachDefaults(store, store.issues as Record<string, unknown>[]);
      break;
    default:
      throw new Error(`Unsupported target`);
  }

  const scope = q.scope;
  if (scope?.kind === "project") {
    rows = rows.filter((r) => String(resolveProject(store, r)) === scope.name);
  }
  if (scope?.kind === "module") {
    rows = rows.filter((r) => String(resolveModule(r)) === scope.name);
  }

  rows = rows.filter((r) => evalWhere(r, q.where));

  if (q.orderBy) {
    const { field, direction } = q.orderBy;
    rows = [...rows].sort((a, b) => compareByField(a, b, field, direction));
  }

  if (q.limit !== undefined) rows = rows.slice(0, q.limit);

  return rows;
}

function attachSemantics(store: IrStore): Record<string, unknown>[] {
  const semantics = store.semantics as Record<string, unknown>[];
  return semantics.map((s) => {
    const origin =
      typeof s.origin === "object" && s.origin !== null ? { ...(s.origin as object) } : {};
    const o = origin as Record<string, unknown>;
    if (o.project === undefined) o.project = store.project;
    return {
      ...s,
      targetKind: "semantic",
      origin: { ...o },
    };
  });
}

function resolveProject(store: IrStore, row: Record<string, unknown>): unknown {
  const fromOrigin = getOrigin(row, "project");
  if (typeof fromOrigin === "string") return fromOrigin;
  if (typeof row.project === "string") return row.project;
  return store.project;
}

function resolveModule(row: Record<string, unknown>): unknown {
  const fromOrigin = getOrigin(row, "module");
  if (typeof fromOrigin === "string") return fromOrigin;
  if (typeof row.module === "string") return row.module;
  const loc = row.location;
  if (typeof loc === "object" && loc !== null) {
    const m = (loc as Record<string, unknown>).module;
    if (typeof m === "string") return m;
  }
  return "";
}

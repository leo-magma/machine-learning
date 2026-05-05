import { readdirSync } from "node:fs";
import { join } from "node:path";

import { crossModuleCount, extractRawOutgoingCalls, resolveOutgoingCalls } from "./callGraph.js";
import { inferSemanticsForProcedure } from "./inferSemantics.js";
import type { IrStore } from "../store.js";
import type { ProcedureScan } from "./scanBas.js";
import {
  estimateCyclomatic,
  hasOnErrorResumeNext,
  listMagicNumberIssues,
  readModuleBas,
} from "./scanBas.js";

function extractSheetsFromSemantic(s: Record<string, unknown>): string[] {
  const out: string[] = [];
  const src = s.source;
  if (typeof src === "object" && src !== null) {
    const sh = (src as Record<string, unknown>).sheet;
    if (typeof sh === "string") out.push(sh);
  }
  for (const side of ["left", "right"]) {
    const o = s[side];
    if (typeof o === "object" && o !== null) {
      const sh = (o as Record<string, unknown>).sheet;
      if (typeof sh === "string") out.push(sh);
    }
  }
  return out;
}

type FlatProc = {
  proc: ProcedureScan;
  moduleName: string;
  mtimeIso: string;
};

/**
 * Scan a folder of exported `.bas` modules and build an in-memory IR store (same shape as JSON `--store`).
 */
export function buildIrStoreFromSamplesDir(samplesDirAbs: string, project: string): IrStore {
  const files = readdirSync(samplesDirAbs).filter((f) => f.toLowerCase().endsWith(".bas"));
  if (files.length === 0) {
    throw new Error(`No .bas files under ${samplesDirAbs}`);
  }

  const flat: FlatProc[] = [];
  for (const file of files.sort()) {
    const abs = join(samplesDirAbs, file);
    const mod = readModuleBas(abs);
    for (const proc of mod.procedures) {
      flat.push({ proc, moduleName: mod.moduleName, mtimeIso: mod.mtimeIso });
    }
  }

  const moduleToProcedures = new Map<string, Set<string>>();
  const procNameToModules = new Map<string, string[]>();
  for (const { proc, moduleName } of flat) {
    if (!moduleToProcedures.has(moduleName)) moduleToProcedures.set(moduleName, new Set());
    moduleToProcedures.get(moduleName)!.add(proc.name);
    const arr = procNameToModules.get(proc.name) ?? [];
    arr.push(moduleName);
    procNameToModules.set(proc.name, arr);
  }

  const semantics: Record<string, unknown>[] = [];
  const procedures: Record<string, unknown>[] = [];
  const issues: Record<string, unknown>[] = [];
  const moduleProcCc: Record<string, number[]> = {};
  const moduleSheetRefs: Record<string, Set<string>> = {};
  const incomingHints = new Map<string, number>();

  const edgeKey = (m: string, p: string): string => `${m}::${p}`;

  for (const { proc, moduleName, mtimeIso } of flat) {
    moduleProcCc[moduleName] ??= [];
    moduleSheetRefs[moduleName] ??= new Set<string>();

    const cc = estimateCyclomatic(proc.body);
    const onErr = hasOnErrorResumeNext(proc.body);
    const magicIss = listMagicNumberIssues(moduleName, proc, project);
    issues.push(...magicIss);

    if (onErr) {
      issues.push({
        kind: "on_error_resume_next",
        severity: "high",
        location: {
          module: moduleName,
          procedure: proc.name,
          lines: { start: proc.startLine, end: proc.startLine },
        },
        origin: { project, module: moduleName, procedure: proc.name },
      });
    }

    const risk = Math.min(
      1,
      Math.round((cc / 40 + (onErr ? 0.28 : 0) + Math.min(0.25, magicIss.length * 0.015)) * 100) /
        100
    );

    const rawCalls = extractRawOutgoingCalls(proc.body);
    const resolvedEdges = resolveOutgoingCalls(moduleName, rawCalls, moduleToProcedures, procNameToModules);

    const outgoingCount = resolvedEdges.length;
    const crossOut = crossModuleCount(moduleName, resolvedEdges);
    const unresolvedOutgoingCount = resolvedEdges.filter((e) => !e.resolved).length;

    for (const e of resolvedEdges) {
      if (!e.resolved) continue;
      const k = edgeKey(e.targetModule, e.targetProcedure);
      incomingHints.set(k, (incomingHints.get(k) ?? 0) + 1);
    }

    appendSemantics(moduleName, proc, project, semantics, moduleSheetRefs);

    procedures.push({
      name: proc.name,
      module: moduleName,
      changedAt: mtimeIso,
      quality: {
        complexity: { cyclomatic: cc },
        risk: {
          riskScore: risk,
          riskScore_delta: Math.round(Math.min(0.5, cc / 100 + risk / 5) * 100) / 100,
        },
        dependency: {
          degree: outgoingCount,
        },
      },
      dependency: {
        outgoing: resolvedEdges.map((e) => ({
          targetModule: e.targetModule,
          targetProcedure: e.targetProcedure,
          resolved: e.resolved,
          raw: e.raw,
        })),
        outgoingCount,
        crossModuleOutgoingCount: crossOut,
        unresolvedOutgoingCount,
        incomingCount: 0,
      },
      origin: {
        project,
        module: moduleName,
        procedure: proc.name,
        lines: { start: proc.startLine, end: proc.endLine },
      },
    });

    moduleProcCc[moduleName]!.push(cc);
  }

  for (const row of procedures) {
    const dep = row.dependency as Record<string, unknown> | undefined;
    if (!dep) continue;
    const mod = row.module as string;
    const name = row.name as string;
    const ic = incomingHints.get(edgeKey(mod, name)) ?? 0;
    dep.incomingCount = ic;
    const q = row.quality as Record<string, unknown>;
    const qualDep = q.dependency as Record<string, unknown>;
    qualDep.degree = Math.round(Number((dep.outgoingCount as number) ?? 0) + ic * 0.5);
  }

  const modules: Record<string, unknown>[] = Object.keys(moduleProcCc).map((name) => {
    const arr = moduleProcCc[name] ?? [];
    const avg = arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 1;

    const outgoingMods = new Set<string>();
    let sheetDeg = moduleSheetRefs[name]?.size ?? 0;
    for (const row of procedures) {
      if ((row.module as string) !== name) continue;
      const dep = row.dependency as { outgoing?: { targetModule: string; resolved: boolean }[] };
      for (const e of dep.outgoing ?? []) {
        if (e.resolved && e.targetModule !== name) outgoingMods.add(e.targetModule);
      }
    }

    const inboundMods = new Set<string>();
    for (const row of procedures) {
      const dep = row.dependency as { outgoing?: { targetModule: string; targetProcedure: string; resolved: boolean }[] };
      const sm = row.module as string;
      if (sm === name) continue;
      for (const e of dep.outgoing ?? []) {
        if (!e.resolved) continue;
        if (e.targetModule === name) inboundMods.add(sm);
      }
    }

    const deg = Math.min(
      40,
      outgoingMods.size * 3 + inboundMods.size * 3 + Math.floor(sheetDeg * 1.2) + Math.floor(arr.length * 1.2)
    );

    return {
      name,
      quality: {
        dependency: { degree: deg },
        complexity: { avgCyclomatic: Math.round(avg * 10) / 10 },
      },
      origin: { project, module: name },
    };
  });

  return {
    project,
    semantics,
    procedures,
    modules,
    issues,
  };
}

function appendSemantics(
  moduleName: string,
  proc: ProcedureScan,
  project: string,
  semantics: Record<string, unknown>[],
  moduleSheetRefs: Record<string, Set<string>>
): void {
  moduleSheetRefs[moduleName] ??= new Set<string>();
  const sem = inferSemanticsForProcedure(project, moduleName, proc);
  for (const s of sem) {
    semantics.push(s);
    const sheets = extractSheetsFromSemantic(s);
    for (const sh of sheets) moduleSheetRefs[moduleName]!.add(sh);
  }
}

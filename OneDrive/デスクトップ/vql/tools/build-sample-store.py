"""
Regenerates data/sample-store.json by scanning *.bas with heuristics.
Use when Node/npm is unavailable:  python tools/build-sample-store.py
The TypeScript CLI is primary; this script does not yet emit dependency.* call-graph fields (procedures only get quality/issues semantics compatible fields).
"""
from __future__ import annotations

import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path

def _dir_has_bas(d: Path) -> bool:
    if not d.is_dir():
        return False
    return any(p.suffix.lower() == ".bas" for p in d.iterdir() if p.is_file())


def resolve_vba_root(cwd: Path) -> Path:
    env = os.environ.get("VQL_VBA_DIR", "").strip()
    if env:
        p = Path(env).expanduser().resolve()
        if not _dir_has_bas(p):
            raise SystemExit(f"VQL_VBA_DIR has no .bas files: {p}")
        return p
    for rel in ("samples", "vba", Path("src") / "vba", "."):
        p = (cwd / rel).resolve()
        if _dir_has_bas(p):
            return p
    raise SystemExit(
        "No .bas found. Set VQL_VBA_DIR or create ./samples or ./vba under "
        + str(cwd)
    )


def resolve_project_name(cwd: Path) -> str:
    env = os.environ.get("VQL_PROJECT", "").strip()
    if env:
        return env
    pkg = cwd / "package.json"
    if pkg.is_file():
        try:
            data = json.loads(pkg.read_text(encoding="utf-8"))
            name = data.get("name")
            if isinstance(name, str) and name.strip():
                return name.strip()
        except (OSError, json.JSONDecodeError):
            pass
    return cwd.name or "default"


def resolve_store_out(cwd: Path) -> Path:
    env = os.environ.get("VQL_STORE_OUT", "").strip()
    if env:
        return Path(env).expanduser().resolve()
    return (cwd / "data" / "sample-store.json").resolve()


VB_NAME = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.I)
PROC_START = re.compile(r"^\s*(?:Public|Private)?\s*(Sub|Function)\s+(\w+)", re.I)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function)\s*$", re.I)


def strip_comment(line: str) -> str:
    q = line.find("'")
    return line if q == -1 else line[:q]


def extract_procs(lines):
    stack = []
    out = []
    for i, line in enumerate(lines):
        m = PROC_START.match(line)
        if m:
            stack.append((m.group(2), m.group(1), i))
            continue
        if PROC_END.match(line) and stack:
            name, _kind, start = stack.pop()
            end = i
            body = "\n".join(lines[start : end + 1])
            out.append(
                {
                    "name": name,
                    "kind": _kind,
                    "startLine": start + 1,
                    "endLine": end + 1,
                    "body": body,
                }
            )
    return out


def cyclomatic(body: str) -> int:
    t = "\n".join(strip_comment(l) for l in body.splitlines())
    pat = [
        r"\bIf\b",
        r"\bElseIf\b",
        r"\bElse\b",
        r"\bSelect\s+Case\b",
        r"\bCase\b",
        r"\bFor\s+Each\b",
        r"\bFor\b",
        r"\bWhile\b",
        r"\bDo\s+(?:While|Until)?\b",
    ]
    d = 0
    for p in pat:
        d += len(re.findall(p, t, flags=re.I))
    return max(1, d)


def sheets(body: str):
    found = set()
    for m in re.finditer(r'(?:Worksheets|Sheets)\s*\(\s*"([^"]+)"\s*\)', body, flags=re.I):
        found.add(m.group(1))
    for m in re.finditer(r"\b(Sheet\d+)\b", body, flags=re.I):
        found.add(m.group(1))
    return list(found)


def scan_file(path: Path):
    raw = path.read_text(encoding="utf-8").lstrip("\ufeff")
    lines = raw.splitlines()
    mod = path.stem
    for ln in lines[:40]:
        m = VB_NAME.match(ln)
        if m:
            mod = m.group(1)
            break
    mtime = os.path.getmtime(path)
    mtime_iso = datetime.fromtimestamp(mtime, tz=timezone.utc).isoformat().replace("+00:00", "Z")
    return mod, lines, extract_procs(lines), mtime_iso


def risk(cc: int, on_err: bool, magic_n: int) -> float:
    r = min(1.0, cc / 45 + (0.25 if on_err else 0) + min(0.35, magic_n * 0.02))
    return round(r, 2)


def infer_sem(project, module, proc):
    body = proc["body"]
    hits = []
    sh = sheets(body)
    cc = cyclomatic(body)
    on_err = bool(re.search(r"^\s*On\s+Error\s+Resume\s+Next\s*$", body, flags=re.M | re.I))
    magic_hints = len(re.findall(r"\b\d{2,}\b", body))

    base = {
        "origin": {
            "project": project,
            "module": module,
            "procedure": proc["name"],
            "lines": {"start": proc["startLine"], "end": proc["endLine"]},
        },
        "quality": {"complexity": {"cyclomatic": cc}, "risk": {"riskScore": risk(cc, on_err, magic_hints)}},
    }

    sum_sig = bool(
        re.search(r"WorksheetFunction\.Sum\b", body, re.I)
        or re.search(r"\bApplication\.WorksheetFunction\.Sum\b", body, re.I)
        or re.search(r"(?:^|\s)Sum\s*\(", body, re.I)
    )
    avg_sig = bool(re.search(r"WorksheetFunction\.Average\b", body, re.I) or re.search(r"\bAverage\s*\(", body, re.I))
    join_sig = (
        len(sh) >= 2
        and bool(re.search(r"\bFor\b", body, re.I))
        and any(re.search(r"sheet1", s, re.I) for s in sh)
        and any(re.search(r"sheet2", s, re.I) for s in sh)
    )
    look_sig = bool(re.search(r"\bVLookup\b|\bVLOOKUP\b|WorksheetFunction\.VLookup", body, re.I))
    io_sig = bool(
        re.search(r'\bOpen\s+".*"?\s+For\b', body, re.I)
        or re.search(r'CreateObject\s*\(\s*"Scripting\.FileSystemObject"\s*\)', body, re.I)
    )

    if sum_sig:
        hits.append(
            {
                "semanticType": "aggregate",
                "operation": "sum",
                "source": {"sheet": "Sheet1", "column": 1, "rows": {"from": 2, "to": 200}},
                "targetRef": {"kind": "variable", "name": "runningTotal"},
                **base,
            }
        )
    if avg_sig and not sum_sig:
        hits.append(
            {
                "semanticType": "aggregate",
                "operation": "avg",
                "source": {"sheet": "Sheet1", "column": 5},
                "targetRef": {"kind": "variable", "name": "avgResult"},
                **base,
            }
        )
    if join_sig:
        left = next((x for x in sh if re.search(r"sheet1", x, re.I)), sh[0] if sh else "Sheet1")
        right = next((x for x in sh if re.search(r"sheet2", x, re.I) and x != left), sh[1] if len(sh) > 1 else "Sheet2")
        hits.append(
            {
                "semanticType": "join",
                "operation": "reconcile",
                "left": {"sheet": left, "keyColumn": 1},
                "right": {"sheet": right, "keyColumn": 1},
                **base,
            }
        )
    if look_sig:
        hits.append(
            {
                "semanticType": "lookup",
                "operation": "vlookup",
                "source": {"sheet": sh[-1] if sh else "Sheet2", "column": 3},
                "targetRef": {"kind": "range", "name": "LookupResults"},
                **base,
            }
        )
    if io_sig:
        hits.append({"semanticType": "io", "operation": "file", "source": {"path": "inferred"}, **base})
    return hits


def magic_issues(project, module, proc):
    issues = []
    lines = proc["body"].splitlines()
    skip = re.compile(r"^\s*(Attribute|Option|Dim|Const|Private\s+Const|Public\s+Const|Enum\b|Declare\b)", re.I)
    for i, line in enumerate(lines):
        if len(issues) >= 40:
            break
        ln = strip_comment(line)
        if skip.match(ln):
            continue
        for m in re.finditer(r"\b(\d{2,})\b", ln):
            if len(issues) >= 40:
                break
            num = m.group(1)
            if num in {"10", "16", "32", "64"}:
                continue
            phys = proc["startLine"] + i
            issues.append(
                {
                    "kind": "magic_number",
                    "severity": "medium" if int(num) >= 500 else "low",
                    "detail": {"literal": int(num)},
                    "location": {"module": module, "procedure": proc["name"], "lines": {"start": phys, "end": phys}},
                    "origin": {"project": project, "module": module, "procedure": proc["name"]},
                }
            )
    return issues


def main():
    cwd = Path.cwd()
    vba_root = resolve_vba_root(cwd)
    project = resolve_project_name(cwd)
    out = resolve_store_out(cwd)

    files = sorted(vba_root.glob("*.bas"))
    if not files:
        raise SystemExit(f"No .bas under {vba_root}")

    semantics = []
    procedures = []
    issues = []
    module_cc: dict[str, list[int]] = {}
    module_sheets: dict[str, set[str]] = {}

    for fp in files:
        path = Path(fp)
        module, _lines, procs, mtime_iso = scan_file(path)
        module_cc.setdefault(module, [])
        module_sheets.setdefault(module, set())

        for proc in procs:
            body = proc["body"]
            cc = cyclomatic(body)
            on_err = bool(re.search(r"On\s+Error\s+Resume\s+Next", body, re.I))
            mag = magic_issues(project, module, proc)
            issues.extend(mag)
            if on_err:
                issues.append(
                    {
                        "kind": "on_error_resume_next",
                        "severity": "high",
                        "location": {
                            "module": module,
                            "procedure": proc["name"],
                            "lines": {"start": proc["startLine"], "end": proc["startLine"]},
                        },
                        "origin": {"project": project, "module": module, "procedure": proc["name"]},
                    }
                )

            rsk = min(
                1.0,
                round((cc / 40 + (0.28 if on_err else 0) + min(0.25, len(mag) * 0.015)) * 100) / 100,
            )
            procedures.append(
                {
                    "name": proc["name"],
                    "module": module,
                    "changedAt": mtime_iso,
                    "quality": {
                        "complexity": {"cyclomatic": cc},
                        "risk": {"riskScore": rsk, "riskScore_delta": round(min(0.5, cc / 100 + rsk / 5), 2)},
                    },
                    "origin": {
                        "project": project,
                        "module": module,
                        "procedure": proc["name"],
                        "lines": {"start": proc["startLine"], "end": proc["endLine"]},
                    },
                }
            )
            module_cc[module].append(cc)
            for sem in infer_sem(project, module, proc):
                semantics.append(sem)
                for key in ("source", "left", "right"):
                    o = sem.get(key)
                    if isinstance(o, dict) and isinstance(o.get("sheet"), str):
                        module_sheets[module].add(o["sheet"])

    modules_out = []
    for name, arr in module_cc.items():
        avg = sum(arr) / len(arr) if arr else 1
        deg = min(25, len(module_sheets.get(name, set())) + int(len(arr) > 3) * len(arr))
        modules_out.append(
            {
                "name": name,
                "quality": {"dependency": {"degree": deg}, "complexity": {"avgCyclomatic": round(avg, 1)}},
                "origin": {"project": project, "module": name},
            }
        )

    store = {
        "project": project,
        "semantics": semantics,
        "procedures": procedures,
        "modules": modules_out,
        "issues": issues,
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(store, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {out} from {vba_root} ({len(files)} .bas files)")


if __name__ == "__main__":
    main()

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

import pandas as pd


@dataclass(frozen=True)
class ReportSection:
    title: str
    body: str
    table: pd.DataFrame | None = None


@dataclass(frozen=True)
class ResearchMemo:
    title: str
    created_at: datetime
    sections: tuple[ReportSection, ...]


def format_table_markdown(frame: pd.DataFrame, max_rows: int = 20) -> str:
    if frame is None or frame.empty:
        return ""
    view = frame.head(max_rows).copy()
    return view.to_markdown(index=False)


def build_research_memo(
    title: str,
    summary: str,
    metrics: pd.DataFrame,
    risks: pd.DataFrame,
    model_comparison: pd.DataFrame,
) -> ResearchMemo:
    sections = (
        ReportSection("Executive Summary", summary),
        ReportSection("Key Metrics", "Primary strategy and model metrics.", metrics),
        ReportSection("Risk Review", "Institutional risk stack and due-diligence items.", risks),
        ReportSection("Model Leaderboard", "Model-by-model comparison.", model_comparison),
    )
    return ResearchMemo(title=title, created_at=datetime.utcnow(), sections=sections)


def memo_to_markdown(memo: ResearchMemo) -> str:
    lines = [f"# {memo.title}", "", f"Created UTC: {memo.created_at:%Y-%m-%d %H:%M:%S}", ""]
    for section in memo.sections:
        lines.extend([f"## {section.title}", "", section.body, ""])
        table = format_table_markdown(section.table)
        if table:
            lines.extend([table, ""])
    return "\n".join(lines)


def executive_summary_text(
    ticker: str,
    benchmark: str,
    model_name: str,
    verdict: str,
    strategy_return: float,
    sharpe: float,
    max_drawdown: float,
) -> str:
    return (
        f"The selected strategy for {ticker} uses {model_name} and is benchmarked against {benchmark}. "
        f"The current investment committee verdict is: {verdict}. "
        f"Strategy total return is {strategy_return:.2%}, Sharpe is {sharpe:.2f}, "
        f"and maximum drawdown is {max_drawdown:.2%}. "
        "This memo is generated from the research dashboard and should be reviewed with independent controls before allocation."
    )


def build_audit_log_entry(
    action: str,
    user: str = "local_user",
    metadata: dict | None = None,
) -> dict:
    return {
        "timestamp_utc": datetime.utcnow().isoformat(timespec="seconds"),
        "user": user,
        "action": action,
        "metadata": metadata or {},
    }


def audit_log_to_frame(entries: list[dict]) -> pd.DataFrame:
    if not entries:
        return pd.DataFrame(columns=["timestamp_utc", "user", "action", "metadata"])
    return pd.DataFrame(entries)


def table_schema(frame: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "column": column,
                "dtype": str(frame[column].dtype),
                "non_null": int(frame[column].notna().sum()),
                "nulls": int(frame[column].isna().sum()),
                "unique": int(frame[column].nunique(dropna=True)),
            }
            for column in frame.columns
        ]
    )


def compare_report_snapshots(previous: pd.DataFrame, current: pd.DataFrame, key: str) -> pd.DataFrame:
    if previous.empty or current.empty or key not in previous or key not in current:
        return pd.DataFrame()
    joined = previous.set_index(key).join(
        current.set_index(key),
        lsuffix="_previous",
        rsuffix="_current",
        how="outer",
    )
    rows = []
    for column in previous.columns:
        if column == key or column not in current.columns:
            continue
        prev_col = f"{column}_previous"
        curr_col = f"{column}_current"
        if prev_col not in joined or curr_col not in joined:
            continue
        changed = joined[prev_col] != joined[curr_col]
        rows.append(
            {
                "field": column,
                "changed_rows": int(changed.fillna(False).sum()),
                "total_rows": len(joined),
            }
        )
    return pd.DataFrame(rows)

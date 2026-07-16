"""
Perzistentní historická data - nahrazují dnešní externí .xlsx mezisoubory
a ruční doplňování Won/Lost. Ukládáme jako CSV (diffovatelné v gitu),
ne jako binární Excel.
"""
from pathlib import Path
from typing import List

import pandas as pd

SNAPSHOTS_COLUMNS = [
    "week_label", "week_monday", "deal_id", "deal_name", "owner_id", "owner_name",
    "stage_id", "stage_label", "amount", "closedate", "pipeline_id",
    "is_closed_won", "is_closed_lost",
]

LEDGER_COLUMNS = [
    "deal_id", "deal_name", "company_name", "owner_name", "closedate",
    "amount", "outcome", "week_label",
]


def load_or_empty(path: Path, columns: List[str]) -> pd.DataFrame:
    if path.exists():
        return pd.read_csv(path)
    return pd.DataFrame(columns=columns)


def append_weekly_snapshot(path: Path, week_label: str, new_rows: pd.DataFrame) -> pd.DataFrame:
    """
    Idempotentní: pokud už pro daný week_label existují řádky, nejdřív se
    smažou a nahradí novými (běh skriptu je bezpečné spustit opakovaně
    ve stejném týdnu, např. při ladění).
    """
    existing = load_or_empty(path, SNAPSHOTS_COLUMNS)
    existing = existing[existing["week_label"] != week_label]
    combined = pd.concat([existing, new_rows], ignore_index=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(path, index=False)
    return combined


def upsert_closed_ledger(path: Path, new_rows: pd.DataFrame) -> pd.DataFrame:
    """Přidá nově uzavřené dealy; podle deal_id nikdy neduplikuje."""
    existing = load_or_empty(path, LEDGER_COLUMNS)
    if not existing.empty:
        known_ids = set(existing["deal_id"].astype(str))
        new_rows = new_rows[~new_rows["deal_id"].astype(str).isin(known_ids)]
    combined = pd.concat([existing, new_rows], ignore_index=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(path, index=False)
    return combined

from pathlib import Path
from typing import List
import pandas as pd

SNAPSHOTS_COLUMNS = ["week_label","week_monday","deal_id","deal_name","owner_id","owner_name",
    "stage_id","stage_label","amount","closedate","pipeline_id","is_closed_won","is_closed_lost"]
LEDGER_COLUMNS = ["deal_id","deal_name","company_name","owner_name","closedate","amount","outcome","week_label"]

def load_or_empty(path, columns):
    if Path(path).exists(): return pd.read_csv(path)
    return pd.DataFrame(columns=columns)

def append_weekly_snapshot(path, week_label, new_rows):
    existing = load_or_empty(path, SNAPSHOTS_COLUMNS)
    existing = existing[existing["week_label"] != week_label]
    combined = pd.concat([existing, new_rows], ignore_index=True)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(path, index=False); return combined

def upsert_closed_ledger(path, new_rows):
    existing = load_or_empty(path, LEDGER_COLUMNS)
    if not existing.empty:
        known = set(existing["deal_id"].astype(str))
        new_rows = new_rows[~new_rows["deal_id"].astype(str).isin(known)]
    combined = pd.concat([existing, new_rows], ignore_index=True)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(path, index=False); return combined

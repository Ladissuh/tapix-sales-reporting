"""
Veškeré odvozené metriky (co dřív dělaly SharePoint formule + tabulky 1/2
na listech per obchodník) - počítáno v pandas z historických dat.
"""
from datetime import datetime, timedelta, date
from typing import Dict, List, Optional

import pandas as pd


def previous_week_label(now_local: datetime):
    monday_this_week = now_local - timedelta(days=now_local.weekday())
    monday_prev = monday_this_week - timedelta(days=7)
    sunday_prev = monday_prev + timedelta(days=6)
    iso = monday_prev.isocalendar()
    label = f"{iso.year}-W{iso.week:02d} ({monday_prev.date()}—{sunday_prev.date()})"
    return label, monday_prev.date(), sunday_prev.date()


def week_label_for_date(d) -> str:
    if isinstance(d, str):
        d = pd.to_datetime(d).date()
    elif isinstance(d, datetime):
        d = d.date()
    monday = d - timedelta(days=d.weekday())
    sunday = monday + timedelta(days=6)
    iso = monday.isocalendar()
    return f"{iso.year}-W{iso.week:02d} ({monday}—{sunday})"


def ordered_week_labels(snapshots_df: pd.DataFrame) -> List[str]:
    if snapshots_df.empty:
        return []
    tmp = snapshots_df[["week_label", "week_monday"]].drop_duplicates()
    tmp["week_monday"] = pd.to_datetime(tmp["week_monday"])
    tmp = tmp.sort_values("week_monday")
    return tmp["week_label"].tolist()


def stage_weighted_amounts_by_owner(
    snapshots_df: pd.DataFrame,
    stage_probabilities: Dict[str, float],
    default_probability: float,
    week_labels: List[str],
    stage_order: List[str],
) -> Dict[str, Dict[str, List[float]]]:
    """owner -> {stage_label: [weighted amount per week in week_labels order]}"""
    result: Dict[str, Dict[str, List[float]]] = {}
    if snapshots_df.empty:
        return result

    grouped = (
        snapshots_df.groupby(["owner_name", "stage_label", "week_label"])["amount"]
        .sum()
        .reset_index()
    )

    for owner in sorted(grouped["owner_name"].unique()):
        result[owner] = {}
        owner_df = grouped[grouped["owner_name"] == owner]
        for stage in stage_order:
            prob = stage_probabilities.get(stage, default_probability)
            stage_df = owner_df[owner_df["stage_label"] == stage]
            by_week = dict(zip(stage_df["week_label"], stage_df["amount"]))
            result[owner][stage] = [round(by_week.get(w, 0.0) * prob) for w in week_labels]
    return result


def raw_amounts_by_owner(
    snapshots_df: pd.DataFrame, week_labels: List[str], stage_order: List[str]
) -> Dict[str, Dict[str, List[float]]]:
    """Stejné jako výše, ale bez váhy - použije se pro funnel graf (nominální Kč)."""
    result: Dict[str, Dict[str, List[float]]] = {}
    if snapshots_df.empty:
        return result
    grouped = (
        snapshots_df.groupby(["owner_name", "stage_label", "week_label"])["amount"]
        .sum()
        .reset_index()
    )
    for owner in sorted(grouped["owner_name"].unique()):
        result[owner] = {}
        owner_df = grouped[grouped["owner_name"] == owner]
        for stage in stage_order:
            stage_df = owner_df[owner_df["stage_label"] == stage]
            by_week = dict(zip(stage_df["week_label"], stage_df["amount"]))
            result[owner][stage] = [round(by_week.get(w, 0.0)) for w in week_labels]
    return result


def won_lost_by_owner(
    ledger_df: pd.DataFrame, week_labels: List[str]
) -> Dict[str, Dict[str, List[float]]]:
    """owner -> {'won': [...], 'lost': [...], 'won_cnt': [...], 'lost_cnt': [...]}"""
    result: Dict[str, Dict[str, List[float]]] = {}
    if ledger_df.empty:
        owners = []
    else:
        owners = sorted(ledger_df["owner_name"].unique())

    for owner in owners:
        odf = ledger_df[ledger_df["owner_name"] == owner]
        won_df = odf[odf["outcome"] == "won"]
        lost_df = odf[odf["outcome"] == "lost"]
        won_by_week = won_df.groupby("close_week_label")["amount"].sum().to_dict()
        lost_by_week = lost_df.groupby("close_week_label")["amount"].sum().to_dict()
        won_cnt_by_week = won_df.groupby("close_week_label")["deal_id"].count().to_dict()
        lost_cnt_by_week = lost_df.groupby("close_week_label")["deal_id"].count().to_dict()
        result[owner] = {
            "won": [round(won_by_week.get(w, 0.0)) for w in week_labels],
            "lost": [round(lost_by_week.get(w, 0.0)) for w in week_labels],
            "won_cnt": [int(won_cnt_by_week.get(w, 0)) for w in week_labels],
            "lost_cnt": [int(lost_cnt_by_week.get(w, 0)) for w in week_labels],
        }
    return result


def pipeline_till_eoy(
    weighted_eoy: Dict[str, List[float]], won: List[float], stage_order: List[str], n: int
) -> List[float]:
    stage_sum = [sum(weighted_eoy[s][w] for s in stage_order) for w in range(n)]
    return [stage_sum[w] + won[w] for w in range(n)]


def week_over_week_changes(values: List[float]) -> List[float]:
    return [values[0]] + [values[w] - values[w - 1] for w in range(1, len(values))]


def build_owner_report_rows(
    weighted_full: Dict[str, List[float]],   # celé 18měs. okno, po stage -> základ pro Rolling 18
    weighted_eoy: Dict[str, List[float]],    # jen do konce roku, po stage -> hlavní tabulka + funnel
    won: List[float],
    lost: List[float],
    won_cnt: List[int],
    lost_cnt: List[int],
    stage_order: List[str],
    annual_goal: Optional[float],
) -> dict:
    """Vrací {'stage_weighted': {...}, 'metrics': {...}, 'annual_goal': ...} - přímo pro report_builder."""
    n = len(won)

    rolling18 = [sum(weighted_full[s][w] for s in stage_order) for w in range(n)]
    pipeline_eoy = pipeline_till_eoy(weighted_eoy, won, stage_order, n)
    changes = week_over_week_changes(pipeline_eoy)

    won_cum, lost_cum = [], []
    wc = lc = 0
    for w in range(n):
        wc += won[w]
        lc += lost[w]
        won_cum.append(wc)
        lost_cum.append(lc)

    win_rate = []
    for w in range(n):
        denom = won_cum[w] + lost_cum[w]
        win_rate.append((won_cum[w] / denom) if denom > 0 else 0.0)

    cum_won_cnt = cum_lost_cnt = 0
    avg_deal = []
    for w in range(n):
        cum_won_cnt += won_cnt[w]
        cum_lost_cnt += lost_cnt[w]
        denom = cum_won_cnt + cum_lost_cnt
        avg_deal.append(((won_cum[w] + lost_cum[w]) / denom) if denom > 0 else 0.0)

    goal_cum = None
    if annual_goal:
        weekly_goal = annual_goal / 52
        goal_cum = [round(weekly_goal * (w + 1)) for w in range(n)]

    metrics = {
        "Won": won,
        "Lost": lost,
        "Pipeline till end of year": pipeline_eoy,
        "Changes in pipeline": changes,
        "Rolling 18": rolling18,
        "Win rate (kumul.)": win_rate,
        "Prům. velikost dealu": avg_deal,
        "Won (kumulativně)": won_cum,
        "Goal (kumulativně)": goal_cum,
    }
    return {"stage_weighted": weighted_eoy, "metrics": metrics, "annual_goal": annual_goal}

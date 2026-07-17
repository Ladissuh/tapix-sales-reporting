from datetime import date, timedelta, datetime
from typing import Dict, List, Optional
import pandas as pd

def previous_week_label(now_local):
    monday_this = now_local - timedelta(days=now_local.weekday())
    monday_prev = monday_this - timedelta(days=7)
    sunday_prev = monday_prev + timedelta(days=6)
    iso = monday_prev.isocalendar()
    return f"{iso.year}-W{iso.week:02d} ({monday_prev.date()}\u2014{sunday_prev.date()})", monday_prev.date(), sunday_prev.date()

def week_label_for_date(d):
    if isinstance(d, str): d = pd.to_datetime(d).date()
    elif isinstance(d, datetime): d = d.date()
    monday = d - timedelta(days=d.weekday())
    sunday = monday + timedelta(days=6)
    iso = monday.isocalendar()
    return f"{iso.year}-W{iso.week:02d} ({monday}\u2014{sunday})"

def ordered_week_labels(df):
    if df.empty: return []
    tmp = df[["week_label","week_monday"]].drop_duplicates()
    tmp = tmp.copy(); tmp["week_monday"] = pd.to_datetime(tmp["week_monday"])
    return tmp.sort_values("week_monday")["week_label"].tolist()

def raw_amounts_by_owner(df, week_labels, stage_order):
    """
    STEJNÉ jako stage_weighted_amounts_by_owner, ale BEZ násobení
    pravděpodobností - přesně to, co stahovaly staré 2 skripty do
    HubSpot_Deals_By_Stage_2026.xlsx / _DYNAMIC. Slouží k debug porovnání.
    """
    result = {}
    if df.empty: return result
    grouped = df.groupby(["owner_name","stage_label","week_label"])["amount"].sum().reset_index()
    for owner in sorted(grouped["owner_name"].unique()):
        result[owner] = {}
        odf = grouped[grouped["owner_name"] == owner]
        for stage in stage_order:
            sdf = odf[odf["stage_label"] == stage]
            by_week = dict(zip(sdf["week_label"], sdf["amount"]))
            result[owner][stage] = [round(by_week.get(w, 0.0)) for w in week_labels]
    return result


def stage_weighted_amounts_by_owner(df, stage_probs, default_prob, week_labels, stage_order):
    result = {}
    if df.empty: return result
    grouped = df.groupby(["owner_name","stage_label","week_label"])["amount"].sum().reset_index()
    for owner in sorted(grouped["owner_name"].unique()):
        result[owner] = {}
        odf = grouped[grouped["owner_name"] == owner]
        for stage in stage_order:
            prob = stage_probs.get(stage, default_prob)
            sdf = odf[odf["stage_label"]==stage]
            by_week = dict(zip(sdf["week_label"], sdf["amount"]))
            result[owner][stage] = [round(by_week.get(w, 0.0)*prob) for w in week_labels]
    return result

def won_lost_by_owner(ledger_df, week_labels):
    result = {}
    if ledger_df.empty: return result
    for owner in sorted(ledger_df["owner_name"].unique()):
        odf = ledger_df[ledger_df["owner_name"]==owner]
        won_df = odf[odf["outcome"]=="won"]; lost_df = odf[odf["outcome"]=="lost"]
        w_bw = won_df.groupby("close_week_label")["amount"].sum().to_dict()
        l_bw = lost_df.groupby("close_week_label")["amount"].sum().to_dict()
        wc_bw = won_df.groupby("close_week_label")["deal_id"].count().to_dict()
        lc_bw = lost_df.groupby("close_week_label")["deal_id"].count().to_dict()
        result[owner] = {
            "won":      [round(w_bw.get(w,0.0)) for w in week_labels],
            "lost":     [round(l_bw.get(w,0.0)) for w in week_labels],
            "won_cnt":  [int(wc_bw.get(w,0))    for w in week_labels],
            "lost_cnt": [int(lc_bw.get(w,0))    for w in week_labels],
        }
    return result

def build_owner_report_rows(weighted_full, weighted_eoy, won, lost, won_cnt, lost_cnt, stage_order, annual_goal,
                             baseline_won=0.0, baseline_lost=0.0, baseline_won_cnt=0, baseline_lost_cnt=0,
                             week_offset=0):
    """
    baseline_* = kumulativní součty ZA VŠECHNY týdny PŘED prvním zobrazeným
    sloupcem (mimo DISPLAY_WEEKS okno) - aby "Won (kumulativně)", "Win rate"
    a "Prům. velikost dealu" odrážely celou historii, ne jen to, co je
    momentálně vidět v tabulce.
    week_offset = kolik týdnů uběhlo PŘED prvním zobrazeným sloupcem - aby
    "Goal (kumulativně)" počítalo se skutečným číslem týdne od začátku roku,
    ne s pořadím v rámci zobrazeného okna.
    """
    n = len(won)
    rolling18 = [sum(weighted_full.get(s,[0]*n)[w] for s in stage_order) for w in range(n)]
    wc = baseline_won; lc = baseline_lost; won_cum = []; lost_cum = []
    for w in range(n):
        wc += won[w]; lc += lost[w]; won_cum.append(wc); lost_cum.append(lc)
    # Pipeline till end of year = vážené stage hodnoty (do konce roku) + KUMULATIVNÍ
    # Won od začátku roku (ne jen týdenní Won) - přesně dle původního vzorce
    # '=(SUM(B55:B65,B67,B70))', kde B67 = "Won stacked" (kumulativní).
    pipeline_eoy = [sum(weighted_eoy.get(s,[0]*n)[w] for s in stage_order)+won_cum[w] for w in range(n)]
    # Changes in Rolling 18 = mezitýdenní delta ROLLING 18 (ne Pipeline till end
    # of year) - přesně dle původního vzorce '=C74-B74', kde řádek 74 = Rolling 18.
    changes_rolling18 = [rolling18[0]]+[rolling18[w]-rolling18[w-1] for w in range(1,n)]
    changes_eoy = [pipeline_eoy[0]]+[pipeline_eoy[w]-pipeline_eoy[w-1] for w in range(1,n)]
    win_rate = [(won_cum[w]/(won_cum[w]+lost_cum[w])) if (won_cum[w]+lost_cum[w]) else 0.0 for w in range(n)]
    cwc = baseline_won_cnt; clc = baseline_lost_cnt; avg_deal = []
    for w in range(n):
        cwc += won_cnt[w]; clc += lost_cnt[w]; d = cwc+clc
        avg_deal.append(((won_cum[w]+lost_cum[w])/d) if d else 0.0)
    goal_cum = [round(annual_goal/52*(week_offset+w+1)) for w in range(n)] if annual_goal else None
    return {"stage_weighted": weighted_eoy, "annual_goal": annual_goal, "metrics": {
        "Won": won, "Lost": lost,
        "Pipeline till end of year": pipeline_eoy,
        "Changes in Rolling 18": changes_rolling18,
        "Changes in pipeline till end of the year": changes_eoy,
        "Rolling 18": rolling18,
        "Win rate (kumul.)": win_rate,
        "Prům. velikost dealu": avg_deal,
        "Won (kumulativně)": won_cum,
        "Goal (kumulativně)": goal_cum,
    }}

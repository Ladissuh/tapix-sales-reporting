#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Týdenní generátor Sales Reporting - jeden skript nahrazující:
  - hubspot_weekly_report_2026.py
  - hubspot_weekly_report_dynamic_2026.py
  - ruční doplňování listů Won/Lost
  - SharePoint Excel formule (Tabulka 1 + Tabulka 2 na listech per obchodník)

Výstup: outputs/Sales_Reporting_<rok>.xlsx - hotový soubor s grafy,
připravený k nahrání na SharePoint (Make.com scénář se teď zjednoduší
na jediný krok: nahradit soubor na SharePointu tímhle výstupem).
"""
import sys
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import yaml
from dateutil.relativedelta import relativedelta

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).parent))

import hubspot_client as hs
import data_store as store
import metrics
import report_builder as rb

LOCAL_TZ = "Europe/Prague"
DISPLAY_WEEKS = 20  # kolik posledních týdnů se zobrazí v reportu

REPO_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = REPO_ROOT / "config"
DATA_DIR = REPO_ROOT / "data"
OUTPUT_DIR = REPO_ROOT / "outputs"

SNAPSHOTS_PATH = DATA_DIR / "deals_snapshots.csv"
LEDGER_PATH = DATA_DIR / "deals_closed_ledger.csv"

DEFAULT_PROBABILITY = 0.5


def load_config():
    with open(CONFIG_DIR / "owners.yaml", encoding="utf-8") as f:
        owners_cfg = yaml.safe_load(f) or {}
    with open(CONFIG_DIR / "stage_probabilities.yaml", encoding="utf-8") as f:
        stage_cfg = yaml.safe_load(f) or {}

    owners_list = owners_cfg.get("owners", []) or []
    # hubspot_name -> display_name (whitelist zároveň slouží jako filtr)
    name_map = {o["hubspot_name"]: o["display_name"] for o in owners_list}
    # display_name -> annual_goal
    goals = {o["display_name"]: o.get("annual_goal") for o in owners_list}
    # pořadí listů v reportu = pořadí v configu
    display_order = [o["display_name"] for o in owners_list]

    stage_probs = stage_cfg.get("stages", {}) or {}
    stage_order = stage_cfg.get("stage_order", list(stage_probs.keys()))
    return name_map, goals, display_order, stage_probs, stage_order


def fetch_and_persist(week_label: str, week_monday: date):
    token = hs.load_token()
    owners_map = hs.get_all_owners(token)
    stage_label_map, _default_order, closed_won_ids, closed_lost_ids = hs.get_stage_metadata(token)

    cutoff_local = datetime.combine(week_monday, datetime.min.time()).replace(
        tzinfo=ZoneInfo(LOCAL_TZ)
    ) + relativedelta(days=6) + relativedelta(months=+18)
    cutoff_ms = int(cutoff_local.timestamp() * 1000)

    deals = hs.fetch_deals(token, cutoff_ms)
    print(f"Staženo {len(deals)} dealů (cutoff {cutoff_local.date()}).")

    open_deals, closed_deals = [], []
    for d in deals:
        stage_id = (d.get("properties") or {}).get("dealstage")
        if stage_id in closed_won_ids or stage_id in closed_lost_ids:
            closed_deals.append(d)
        else:
            open_deals.append(d)

    deal_ids = [d.get("id") for d in closed_deals]
    company_ids_by_deal = hs.get_deal_company_ids(token, deal_ids)
    company_names = hs.get_company_names(token, list(company_ids_by_deal.values()))

    # ---- Weekly snapshot (jen otevřené dealy) ----
    snap_rows = []
    for d in open_deals:
        props = d.get("properties") or {}
        stage_id = props.get("dealstage")
        owner_id = props.get("hubspot_owner_id")
        amount = props.get("amount")
        try:
            amount_val = float(amount) if amount not in (None, "") else 0.0
        except ValueError:
            amount_val = 0.0
        snap_rows.append({
            "week_label": week_label,
            "week_monday": week_monday.isoformat(),
            "deal_id": d.get("id"),
            "deal_name": props.get("dealname", ""),
            "owner_id": owner_id,
            "owner_name": owners_map.get(str(owner_id), "Unassigned"),
            "stage_id": stage_id,
            "stage_label": stage_label_map.get(stage_id, "Unknown stage"),
            "amount": amount_val,
            "closedate": props.get("closedate"),
            "pipeline_id": props.get("pipeline"),
            "is_closed_won": False,
            "is_closed_lost": False,
        })
    snap_df = pd.DataFrame(snap_rows)
    store.append_weekly_snapshot(SNAPSHOTS_PATH, week_label, snap_df)

    # ---- Closed ledger (Won/Lost, upsert podle deal_id) ----
    ledger_rows = []
    for d in closed_deals:
        props = d.get("properties") or {}
        stage_id = props.get("dealstage")
        owner_id = props.get("hubspot_owner_id")
        amount = props.get("amount")
        try:
            amount_val = float(amount) if amount not in (None, "") else 0.0
        except ValueError:
            amount_val = 0.0
        cid = company_ids_by_deal.get(str(d.get("id")), "")
        closedate_raw = props.get("closedate")
        close_week_label = metrics.week_label_for_date(closedate_raw) if closedate_raw else week_label
        ledger_rows.append({
            "deal_id": d.get("id"),
            "deal_name": props.get("dealname", ""),
            "company_name": company_names.get(cid, ""),
            "owner_name": owners_map.get(str(owner_id), "Unassigned"),
            "closedate": closedate_raw,
            "amount": amount_val,
            "outcome": "won" if stage_id in closed_won_ids else "lost",
            "week_label": close_week_label,
        })
    ledger_df = pd.DataFrame(ledger_rows)
    store.upsert_closed_ledger(LEDGER_PATH, ledger_df)


def build_report(week_monday: date):
    name_map, goals, display_order, stage_probs, stage_order = load_config()

    snapshots_df = store.load_or_empty(SNAPSHOTS_PATH, store.SNAPSHOTS_COLUMNS)
    ledger_df = store.load_or_empty(LEDGER_PATH, store.LEDGER_COLUMNS)

    # Whitelist + přejmenování na display_name - VŠECHNO ostatní (bývalí
    # zaměstnanci, kolegové z jiných týmů, Unassigned...) se odsud vypadne
    # a dál v reportu vůbec neexistuje.
    if not snapshots_df.empty:
        snapshots_df = snapshots_df[snapshots_df["owner_name"].isin(name_map.keys())].copy()
        snapshots_df["owner_name"] = snapshots_df["owner_name"].map(name_map)
    if not ledger_df.empty:
        ledger_df = ledger_df[ledger_df["owner_name"].isin(name_map.keys())].copy()
        ledger_df["owner_name"] = ledger_df["owner_name"].map(name_map)
        ledger_df = ledger_df.rename(columns={"week_label": "close_week_label"})

    all_week_labels = metrics.ordered_week_labels(snapshots_df)
    week_labels = all_week_labels[-DISPLAY_WEEKS:] if len(all_week_labels) > DISPLAY_WEEKS else all_week_labels
    if not week_labels:
        raise RuntimeError("Žádná historická data pro nakonfigurované obchodníky - "
                            "zkontroluj, jestli hubspot_name v config/owners.yaml přesně sedí.")

    # Krátký, čitelný label do hlavičky sloupců (jen datum pondělí)
    week_monday_lookup = (
        snapshots_df[["week_label", "week_monday"]].drop_duplicates().set_index("week_label")["week_monday"]
    )
    week_labels_display = []
    for w in week_labels:
        try:
            d = pd.to_datetime(week_monday_lookup[w]).date()
            week_labels_display.append(d.strftime("%d.%m."))
        except KeyError:
            week_labels_display.append(w)

    eoy_cutoff = date(week_monday.year + 1, 1, 1)
    snapshots_df_dates = pd.to_datetime(snapshots_df["closedate"], errors="coerce").dt.date if not snapshots_df.empty else None
    if snapshots_df.empty:
        snapshots_eoy_df = snapshots_df
    else:
        snapshots_eoy_df = snapshots_df[snapshots_df_dates < eoy_cutoff]

    weighted_full = metrics.stage_weighted_amounts_by_owner(
        snapshots_df, stage_probs, DEFAULT_PROBABILITY, week_labels, stage_order
    )
    weighted_eoy = metrics.stage_weighted_amounts_by_owner(
        snapshots_eoy_df, stage_probs, DEFAULT_PROBABILITY, week_labels, stage_order
    )
    wl = metrics.won_lost_by_owner(ledger_df, week_labels)

    # Pevné pořadí a přesně ti lidé, co jsou v config/owners.yaml - ne "kdokoliv
    # se objevil v datech".
    owners_present = display_order

    wb = rb.new_workbook()

    # Won / Lost ledger sheets
    def ledger_rows_for(outcome):
        if ledger_df.empty:
            return []
        odf = ledger_df[ledger_df["outcome"] == outcome]
        rows = []
        for _, row in odf.iterrows():
            try:
                cd = pd.to_datetime(row["closedate"]).date()
            except Exception:
                cd = None
            rows.append([row["deal_id"], row["deal_name"], row["company_name"], row["owner_name"],
                         cd, row["amount"], ""])
        return rows

    rb.build_ledger_sheet(wb, "Won", ledger_rows_for("won"), rb.GREEN)
    rb.build_ledger_sheet(wb, "Lost", ledger_rows_for("lost"), rb.RED)

    leaderboard_totals = {}
    for owner in owners_present:
        wl_owner = wl.get(owner, {"won": [0] * len(week_labels), "lost": [0] * len(week_labels),
                                   "won_cnt": [0] * len(week_labels), "lost_cnt": [0] * len(week_labels)})
        w_full = weighted_full.get(owner, {s: [0] * len(week_labels) for s in stage_order})
        w_eoy = weighted_eoy.get(owner, {s: [0] * len(week_labels) for s in stage_order})
        annual_goal = goals.get(owner)

        rows = metrics.build_owner_report_rows(
            weighted_full=w_full, weighted_eoy=w_eoy,
            won=wl_owner["won"], lost=wl_owner["lost"],
            won_cnt=wl_owner["won_cnt"], lost_cnt=wl_owner["lost_cnt"],
            stage_order=stage_order, annual_goal=annual_goal,
        )
        rb.build_person_sheet(wb, owner, rows, stage_order, week_labels_display, week_labels_display)
        leaderboard_totals[owner] = sum(wl_owner["won"])

    # Aggregation
    agg_stage_full = {s: [sum(weighted_full.get(o, {}).get(s, [0] * len(week_labels))[w] for o in owners_present)
                           for w in range(len(week_labels))] for s in stage_order}
    agg_stage_eoy = {s: [sum(weighted_eoy.get(o, {}).get(s, [0] * len(week_labels))[w] for o in owners_present)
                          for w in range(len(week_labels))] for s in stage_order}
    agg_won = [sum(wl.get(o, {"won": [0] * len(week_labels)})["won"][w] for o in owners_present) for w in range(len(week_labels))]
    agg_lost = [sum(wl.get(o, {"lost": [0] * len(week_labels)})["lost"][w] for o in owners_present) for w in range(len(week_labels))]
    agg_won_cnt = [sum(wl.get(o, {"won_cnt": [0] * len(week_labels)})["won_cnt"][w] for o in owners_present) for w in range(len(week_labels))]
    agg_lost_cnt = [sum(wl.get(o, {"lost_cnt": [0] * len(week_labels)})["lost_cnt"][w] for o in owners_present) for w in range(len(week_labels))]
    total_goal = sum(v for v in goals.values() if v)

    agg_rows = metrics.build_owner_report_rows(
        weighted_full=agg_stage_full, weighted_eoy=agg_stage_eoy,
        won=agg_won, lost=agg_lost, won_cnt=agg_won_cnt, lost_cnt=agg_lost_cnt,
        stage_order=stage_order, annual_goal=total_goal or None,
    )
    leaderboard = sorted(leaderboard_totals.items(), key=lambda x: -x[1])
    rb.build_aggregation_sheet(wb, agg_rows, stage_order, week_labels_display, week_labels_display, leaderboard)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUTPUT_DIR / f"Sales_Reporting_{week_monday.year}.xlsx"
    wb.save(out_path)
    print(f"Report uložen: {out_path}")
    return out_path


def main():
    now_local = datetime.now(ZoneInfo(LOCAL_TZ))
    week_label, week_monday, _sunday = metrics.previous_week_label(now_local)
    print(f"Zpracovávám týden: {week_label}")
    fetch_and_persist(week_label, week_monday)
    build_report(week_monday)


if __name__ == "__main__":
    main()

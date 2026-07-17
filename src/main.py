#!/usr/bin/env python3
"""Týdenní generátor Sales Report z HubSpotu."""
import sys
from datetime import datetime, date, timedelta
from pathlib import Path
import pandas as pd, yaml
from dateutil.relativedelta import relativedelta
try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).parent))
import hubspot_client as hs, data_store as store, metrics, report_builder as rb

LOCAL_TZ = "Europe/Prague"
DISPLAY_WEEKS = 20
REPO_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = REPO_ROOT / "config"
DATA_DIR   = REPO_ROOT / "data"
OUTPUT_DIR = REPO_ROOT / "outputs"
SNAPSHOTS_PATH = DATA_DIR / "deals_snapshots.csv"
LEDGER_PATH    = DATA_DIR / "deals_closed_ledger.csv"
DEFAULT_PROBABILITY = 0.5

# Stage labely které považujeme za uzavřené (záloha pokud probability metadata chybí)
CLOSED_WON_LABELS  = {"closed won", "won", "closed"}
CLOSED_LOST_LABELS = {"closed lost", "lost", "lost (tapix)"}

def load_config():
    with open(CONFIG_DIR / "owners.yaml", encoding="utf-8") as f:
        oc = yaml.safe_load(f) or {}
    with open(CONFIG_DIR / "stage_probabilities.yaml", encoding="utf-8") as f:
        sc = yaml.safe_load(f) or {}
    ol = oc.get("owners", []) or []
    name_map      = {o["hubspot_name"]: o["display_name"] for o in ol}
    goals         = {o["display_name"]: o.get("annual_goal") for o in ol}
    display_order = [o["display_name"] for o in ol]
    stage_probs   = sc.get("stages", {}) or {}
    stage_order   = sc.get("stage_order", list(stage_probs.keys()))
    return name_map, goals, display_order, stage_probs, stage_order

def is_closed_won(stage_id, stage_label, closed_won_ids):
    if stage_id in closed_won_ids:
        return True
    return (stage_label or "").strip().lower() in CLOSED_WON_LABELS

def is_closed_lost(stage_id, stage_label, closed_lost_ids):
    if stage_id in closed_lost_ids:
        return True
    return (stage_label or "").strip().lower() in CLOSED_LOST_LABELS

def fetch_and_persist(week_label, week_monday):
    token = hs.load_token()
    owners_map = hs.get_all_owners(token)
    stage_label_map, _, closed_won_ids, closed_lost_ids = hs.get_stage_metadata(token)

    print(f"Closed Won stage IDs: {closed_won_ids}")
    print(f"Closed Lost stage IDs: {closed_lost_ids}")

    cutoff_local = datetime.combine(week_monday, datetime.min.time()).replace(
        tzinfo=ZoneInfo(LOCAL_TZ)) + timedelta(days=6) + relativedelta(months=+18)
    cutoff_ms = int(cutoff_local.timestamp() * 1000)
    deals = hs.fetch_deals(token, cutoff_ms)
    print(f"Staženo celkem {len(deals)} dealů.")

    open_deals, closed_deals = [], []
    for d in deals:
        p   = d.get("properties") or {}
        sid = p.get("dealstage")
        lbl = stage_label_map.get(sid, "")
        if is_closed_won(sid, lbl, closed_won_ids) or is_closed_lost(sid, lbl, closed_lost_ids):
            closed_deals.append(d)
        else:
            open_deals.append(d)

    print(f"  Otevřené dealy (půjdou do pipeline snapshotu): {len(open_deals)}")
    print(f"  Uzavřené dealy (Won/Lost ledger):              {len(closed_deals)}")

    # Debug: ukáž stage rozložení otevřených dealů
    from collections import Counter
    stage_counts = Counter(stage_label_map.get((d.get("properties") or {}).get("dealstage"),"?") for d in open_deals)
    print(f"  Stage breakdown: {dict(stage_counts)}")

    # Batch company lookup pro uzavřené dealy
    deal_ids    = [d.get("id") for d in closed_deals]
    cid_by_deal = hs.get_deal_company_ids(token, deal_ids)
    cnames      = hs.get_company_names(token, list(set(cid_by_deal.values())))

    # Snapshot z otevřených dealů
    snap_rows = []
    for d in open_deals:
        p = d.get("properties") or {}
        sid = p.get("dealstage"); oid = p.get("hubspot_owner_id")
        try: amt = float(p.get("amount") or 0)
        except ValueError: amt = 0.0
        snap_rows.append({
            "week_label":    week_label,
            "week_monday":   week_monday.isoformat(),
            "deal_id":       d.get("id"),
            "deal_name":     p.get("dealname", ""),
            "owner_id":      oid,
            "owner_name":    owners_map.get(str(oid), "Unassigned"),
            "stage_id":      sid,
            "stage_label":   stage_label_map.get(sid, "Unknown stage"),
            "amount":        amt,
            "closedate":     p.get("closedate", ""),
            "pipeline_id":   p.get("pipeline", ""),
            "is_closed_won": False,
            "is_closed_lost":False,
        })
    store.append_weekly_snapshot(SNAPSHOTS_PATH, week_label, pd.DataFrame(snap_rows))
    print(f"  Snapshot uložen: {len(snap_rows)} řádků pro {week_label}")

    # Won/Lost ledger
    ledger_rows = []
    for d in closed_deals:
        p   = d.get("properties") or {}
        sid = p.get("dealstage"); oid = p.get("hubspot_owner_id")
        lbl = stage_label_map.get(sid, "")
        try: amt = float(p.get("amount") or 0)
        except ValueError: amt = 0.0
        cid    = cid_by_deal.get(str(d.get("id")), "")
        cd_raw = p.get("closedate", "")
        close_wl = metrics.week_label_for_date(cd_raw) if cd_raw else week_label
        ledger_rows.append({
            "deal_id":      d.get("id"),
            "deal_name":    p.get("dealname", ""),
            "company_name": cnames.get(cid, ""),
            "owner_name":   owners_map.get(str(oid), "Unassigned"),
            "closedate":    cd_raw,
            "amount":       amt,
            "outcome":      "won" if is_closed_won(sid, lbl, closed_won_ids) else "lost",
            "week_label":   close_wl,
        })
    store.upsert_closed_ledger(LEDGER_PATH, pd.DataFrame(ledger_rows))
    print(f"  Ledger aktualizován: {len(ledger_rows)} nových/existujících záznamů")

def build_report(week_monday):
    name_map, goals, display_order, stage_probs, stage_order = load_config()
    snap   = store.load_or_empty(SNAPSHOTS_PATH, store.SNAPSHOTS_COLUMNS)
    ledger = store.load_or_empty(LEDGER_PATH,    store.LEDGER_COLUMNS)

    if not snap.empty:
        snap = snap[snap["owner_name"].isin(name_map)].copy()
        snap["owner_name"] = snap["owner_name"].map(name_map)
    if not ledger.empty:
        ledger = ledger[ledger["owner_name"].isin(name_map)].copy()
        ledger["owner_name"] = ledger["owner_name"].map(name_map)
        ledger = ledger.rename(columns={"week_label": "close_week_label"})

    all_wl      = metrics.ordered_week_labels(snap)
    week_labels = all_wl[-DISPLAY_WEEKS:] if len(all_wl) > DISPLAY_WEEKS else all_wl
    if not week_labels:
        raise RuntimeError("Žádná historická data.")

    wml = snap[["week_label","week_monday"]].drop_duplicates().set_index("week_label")["week_monday"]
    week_labels_display = []
    for w in week_labels:
        try:
            d = pd.to_datetime(wml[w]).date()
            week_labels_display.append((d + timedelta(days=6)).strftime("%d.%m."))
        except KeyError:
            week_labels_display.append(w)

    eoy      = date(week_monday.year + 1, 1, 1)
    snap_eoy = snap[pd.to_datetime(snap["closedate"], errors="coerce").dt.date < eoy] if not snap.empty else snap

    wf  = metrics.stage_weighted_amounts_by_owner(snap,     stage_probs, DEFAULT_PROBABILITY, week_labels, stage_order)
    we  = metrics.stage_weighted_amounts_by_owner(snap_eoy, stage_probs, DEFAULT_PROBABILITY, week_labels, stage_order)
    wl  = metrics.won_lost_by_owner(ledger, week_labels)

    owners = display_order
    z  = [0] * len(week_labels)
    wb = rb.new_workbook()

    def ledger_rows_for(outcome):
        if ledger.empty: return []
        odf = ledger[ledger["outcome"] == outcome]; rows = []
        for _, row in odf.iterrows():
            try: cd = pd.to_datetime(row["closedate"]).date()
            except: cd = None
            rows.append([row["deal_id"], row["deal_name"], row["company_name"],
                         row["owner_name"], cd, row["amount"], ""])
        return rows

    rb.build_ledger_sheet(wb, "Won",  ledger_rows_for("won"),  rb.GREEN)
    rb.build_ledger_sheet(wb, "Lost", ledger_rows_for("lost"), rb.RED)

    lb_totals = {}
    for owner in owners:
        wl_o = wl.get(owner, {"won": z, "lost": z, "won_cnt": z, "lost_cnt": z})
        rows = metrics.build_owner_report_rows(
            wf.get(owner, {s: z for s in stage_order}),
            we.get(owner, {s: z for s in stage_order}),
            wl_o["won"], wl_o["lost"], wl_o["won_cnt"], wl_o["lost_cnt"],
            stage_order, goals.get(owner))
        rb.build_person_sheet(wb, owner, rows, stage_order, week_labels_display, week_labels_display)
        lb_totals[owner] = sum(wl_o["won"])

    agg_wf = {s: [sum(wf.get(o,{}).get(s,z)[w] for o in owners) for w in range(len(week_labels))] for s in stage_order}
    agg_we = {s: [sum(we.get(o,{}).get(s,z)[w] for o in owners) for w in range(len(week_labels))] for s in stage_order}
    agg_won  = [sum(wl.get(o,{"won":  z})["won"][w]      for o in owners) for w in range(len(week_labels))]
    agg_lost = [sum(wl.get(o,{"lost": z})["lost"][w]     for o in owners) for w in range(len(week_labels))]
    agg_wc   = [sum(wl.get(o,{"won_cnt": z})["won_cnt"][w]   for o in owners) for w in range(len(week_labels))]
    agg_lc   = [sum(wl.get(o,{"lost_cnt":z})["lost_cnt"][w]  for o in owners) for w in range(len(week_labels))]
    total_goal = sum(v for v in goals.values() if v)
    agg_rows = metrics.build_owner_report_rows(
        agg_wf, agg_we, agg_won, agg_lost, agg_wc, agg_lc,
        stage_order, total_goal or None)
    rb.build_aggregation_sheet(wb, agg_rows, stage_order, week_labels_display, week_labels_display,
                               sorted(lb_totals.items(), key=lambda x: -x[1]))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / f"Sales_Reporting_{week_monday.year}.xlsx"
    wb.save(out); print(f"Report uložen: {out}"); return out

def main():
    now = datetime.now(ZoneInfo(LOCAL_TZ))
    week_label, week_monday, _ = metrics.previous_week_label(now)
    print(f"Zpracovávám: {week_label}")
    fetch_and_persist(week_label, week_monday)
    build_report(week_monday)

if __name__ == "__main__":
    main()

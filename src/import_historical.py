#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JEDNORÁZOVÝ import historie z dosavadního Sales_reporting_2026.xlsx do nového
CSV formátu (data/deals_snapshots.csv, data/deals_closed_ledger.csv).

Spusť JEDNOU, než necháš GitHub Actions workflow běžet poprvé naostro. Pak už
skript main.py jen přidává nové týdny na tohle jako základ.

Použití:
    python src/import_historical.py /cesta/k/Sales_reporting_2026.xlsx

Company název u historických Won/Lost řádků se vyplní jen pokud je nastavený
HUBSPOT_TOKEN (viz .env) - jinak zůstane prázdný (do budoucna se stejně bude
plnit automaticky).
"""
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
import metrics
import data_store as store
import hubspot_client as hs

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data"

# Mapování NÁZEV LISTU (v dosavadním excelu) -> PŘESNÉ jméno z HubSpotu
# (musí sedět s "hubspot_name" v config/owners.yaml).
SHEET_TO_HUBSPOT_NAME = {
    "Martin": "Martin Korbelar",
    "Lukáš": "Lukas Hora",
    "Veronika": "Veronika Kincová",
    "Mykyta": "Mykyta Artiukhov",
    "Richard": "Richard Brůža",
    "Olda": "Oldřich Huzil",
    "Julie": "Julie Mrkvickova",
    # "Ainaz" záměrně vynechána - už ve firmě není (viz zadání).
}

STAGE_ROW_LABELS = [
    "Lead Engaged", "Meeting negotiation / contacted", "Intro meeting agreed",
    "Awaiting confirmation / Not Now", "Qualify", "Discover", "Validate",
    "Decide", "Commit", "Tech intro", "Implementation", "Testing",
]


def import_stage_snapshots(wb) -> pd.DataFrame:
    rows = []
    for sheet_name, hubspot_name in SHEET_TO_HUBSPOT_NAME.items():
        if sheet_name not in wb.sheetnames:
            print(f"  přeskakuji {sheet_name} - list v souboru neexistuje")
            continue
        ws = wb[sheet_name]

        # Najdi poslední týden se skutečnými daty (Sum řádek = 16),
        # dokud nenarazíme na #N/A / prázdno.
        last_valid_col = 1
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=16, column=c).value
            if isinstance(v, (int, float)):
                last_valid_col = c
            else:
                break

        # Řádky 4-15 = jednotlivé stage (label ve sloupci A, může mít mezery navíc)
        stage_row_by_label = {}
        for r in range(4, 16):
            label = ws.cell(row=r, column=1).value
            if label:
                stage_row_by_label[label.strip()] = r

        n_weeks_imported = 0
        for c in range(2, last_valid_col + 1):
            week_date = ws.cell(row=3, column=c).value
            if not isinstance(week_date, datetime):
                continue
            week_date = week_date.date()
            week_label = metrics.week_label_for_date(week_date)
            n_weeks_imported += 1

            for stage_label in STAGE_ROW_LABELS:
                r = stage_row_by_label.get(stage_label)
                if r is None:
                    continue
                amount = ws.cell(row=r, column=c).value
                if not isinstance(amount, (int, float)):
                    amount = 0
                rows.append({
                    "week_label": week_label,
                    "week_monday": week_date.isoformat(),
                    "deal_id": f"hist-{sheet_name}-{stage_label}-{c}".replace(" ", "_"),
                    "deal_name": "",
                    "owner_id": "",
                    "owner_name": hubspot_name,
                    "stage_id": stage_label,
                    "stage_label": stage_label,
                    "amount": float(amount),
                    "closedate": week_date.isoformat(),  # viz poznámka v README importu
                    "pipeline_id": "",
                    "is_closed_won": False,
                    "is_closed_lost": False,
                })
        print(f"  {sheet_name} ({hubspot_name}): {n_weeks_imported} týdnů importováno")

    return pd.DataFrame(rows, columns=store.SNAPSHOTS_COLUMNS)


def import_closed_ledger(wb, token: str = None) -> pd.DataFrame:
    company_names = {}
    company_ids_needed = set()

    raw_rows = []
    for sheet_name, outcome in [("Won", "won"), ("Lost", "lost")]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for r in range(2, ws.max_row + 1):
            close_date = ws.cell(row=r, column=1).value
            amount = ws.cell(row=r, column=4).value
            owner_name = ws.cell(row=r, column=5).value
            deal_id = ws.cell(row=r, column=7).value
            company_id = ws.cell(row=r, column=8).value
            if not deal_id or not owner_name:
                continue
            if isinstance(close_date, datetime):
                close_date_iso = close_date.isoformat()
                close_week_label = metrics.week_label_for_date(close_date.date())
            else:
                close_date_iso = ""
                close_week_label = ""
            raw_rows.append({
                "deal_id": str(deal_id),
                "deal_name": "",
                "company_id": str(company_id) if company_id else "",
                "owner_name": owner_name,
                "closedate": close_date_iso,
                "amount": float(amount) if isinstance(amount, (int, float)) else 0.0,
                "outcome": outcome,
                "week_label": close_week_label,
            })
            if company_id:
                company_ids_needed.add(str(company_id))
        print(f"  {sheet_name}: {sum(1 for r in raw_rows if r['outcome'] == outcome)} řádků")

    if token and company_ids_needed:
        print(f"  Dotahuji {len(company_ids_needed)} názvů firem z HubSpotu...")
        company_names = hs.get_company_names(token, list(company_ids_needed))

    for r in raw_rows:
        r["company_name"] = company_names.get(r.pop("company_id"), "") if "company_id" in r else r.get("company_name", "")
    # (company_id byl dočasný klíč - výše ho nahrazujeme za company_name)

    return pd.DataFrame(raw_rows, columns=store.LEDGER_COLUMNS)


def main():
    if len(sys.argv) < 2:
        print("Použití: python src/import_historical.py /cesta/k/Sales_reporting_2026.xlsx")
        sys.exit(1)

    xlsx_path = Path(sys.argv[1])
    print(f"Načítám {xlsx_path}...")
    wb = load_workbook(xlsx_path, data_only=True)

    token = None
    try:
        token = hs.load_token()
    except RuntimeError:
        print("  (HUBSPOT_TOKEN nenalezen - Company název u historických Won/Lost zůstane prázdný)")

    print("Importuji stage snapshoty (Tabulka 1 z listů per obchodník)...")
    snapshots_df = import_stage_snapshots(wb)

    print("Importuji Won/Lost ledger...")
    ledger_df = import_closed_ledger(wb, token)

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    snapshots_path = DATA_DIR / "deals_snapshots.csv"
    ledger_path = DATA_DIR / "deals_closed_ledger.csv"

    if snapshots_path.exists():
        print(f"POZOR: {snapshots_path} už existuje - přepisuji.")
    if ledger_path.exists():
        print(f"POZOR: {ledger_path} už existuje - přepisuji.")

    snapshots_df.to_csv(snapshots_path, index=False)
    ledger_df.to_csv(ledger_path, index=False)

    print(f"\nHotovo:")
    print(f"  {snapshots_path} - {len(snapshots_df)} řádků")
    print(f"  {ledger_path} - {len(ledger_df)} řádků")
    print("\nTyhle 2 soubory teď commitni do repa (do složky data/) - "
          "GitHub Actions na ně bude od příštího běhu jen přidávat nové týdny.")


if __name__ == "__main__":
    main()

from pathlib import Path
import pandas as pd

SNAPSHOTS_COLUMNS = ["week_label","week_monday","deal_id","deal_name","owner_id","owner_name",
    "stage_id","stage_label","amount","closedate","pipeline_id","is_closed_won","is_closed_lost"]
LEDGER_COLUMNS = ["deal_id","deal_name","company_name","owner_name","closedate","amount","outcome","week_label"]

def load_or_empty(path, columns):
    if Path(path).exists(): return pd.read_csv(path)
    return pd.DataFrame(columns=columns)

def append_weekly_snapshot(path, week_label, new_rows, force=False):
    """
    Idempotentní zápis týdenního snapshotu, ALE s ochranou proti tichému
    přepsání již UZAVŘENÉHO (minulého) týdne.

    - Týden, který je stále "nejnovější" (žádný pozdější týden ještě v datech
      není), se dál smí libovolně přepisovat - to je záměrné a potřebné
      (idempotentní re-run, když první fetch selže/je nekompletní).
    - Jakmile je ale v datech zaznamenán NOVĚJŠÍ týden, všechny starší týdny
      se považují za uzavřené/zamčené - další pokus o jejich přepsání skončí
      chybou, ne tichou změnou historie. Přesně tohle způsobilo, že se čísla
      z minulého týdne jednou "sama" změnila.
    - force=True tuhle pojistku obchází (např. pro ruční jednorázovou opravu
      historie, jako je doplnění správných dat za W28).
    """
    existing = load_or_empty(path, SNAPSHOTS_COLUMNS)
    if not existing.empty and not force and not new_rows.empty:
        this_monday = pd.to_datetime(new_rows["week_monday"].iloc[0]).date()
        other_weeks = existing[existing["week_label"] != week_label]
        if not other_weeks.empty:
            max_existing_monday = pd.to_datetime(other_weeks["week_monday"]).max().date()
            if this_monday < max_existing_monday:
                raise RuntimeError(
                    f"ZAMČENÝ TÝDEN: {week_label} (týden začínající {this_monday}) je "
                    f"starší než nejnovější již uložený týden ({max_existing_monday}). "
                    "Jakmile je zaznamenán novější týden, starší týdny se považují za "
                    "uzavřené a chráněné proti přepsání - jinak by se mohla čísla z "
                    "minulých týdnů kdykoliv tiše změnit. Pokud tento týden opravdu "
                    "chceš přepsat (např. jednorázová oprava chybného fetchu), zavolej "
                    "append_weekly_snapshot(..., force=True)."
                )
    # Vždy smaž stávající záznamy pro daný týden a nahraď novými
    existing = existing[existing["week_label"] != week_label]
    combined = pd.concat([existing, new_rows], ignore_index=True)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(path, index=False)
    return combined

def upsert_closed_ledger(path, new_rows):
    existing = load_or_empty(path, LEDGER_COLUMNS)
    if not existing.empty:
        known = set(existing["deal_id"].astype(str))
        new_rows = new_rows[~new_rows["deal_id"].astype(str).isin(known)]
    combined = pd.concat([existing, new_rows], ignore_index=True)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(path, index=False)
    return combined

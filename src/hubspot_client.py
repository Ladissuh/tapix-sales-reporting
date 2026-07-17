"""Veškerá komunikace s HubSpot API."""
import os, time
from typing import Dict, List, Tuple
from pathlib import Path
import requests
from dotenv import load_dotenv

HUBSPOT_BASE = "https://api.hubapi.com"

def load_token():
    load_dotenv(dotenv_path=Path(__file__).parent.parent / ".env")
    token = os.getenv("HUBSPOT_TOKEN")
    if not token: raise RuntimeError("Chybí HUBSPOT_TOKEN.")
    return token

def _h(token): return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
def _sleep(a): time.sleep(min(2**a, 32))

def _get(url, token, params=None):
    a = 0
    while True:
        r = requests.get(url, headers=_h(token), params=params)
        if r.status_code == 429 or 500 <= r.status_code < 600: a+=1; _sleep(a); continue
        r.raise_for_status(); return r.json()

def _post(url, token, body):
    a = 0
    while True:
        r = requests.post(url, headers=_h(token), json=body)
        if r.status_code == 429 or 500 <= r.status_code < 600: a+=1; _sleep(a); continue
        if r.status_code >= 400:
            print("HubSpot API error body:", r.text[:1000])
        r.raise_for_status(); return r.json()

def get_all_owners(token):
    out = {}; params = {"limit": 100, "archived": "false"}
    while True:
        data = _get(f"{HUBSPOT_BASE}/crm/v3/owners/", token, params)
        for o in data.get("results", []):
            oid = str(o.get("id"))
            out[oid] = f"{o.get('firstName','')} {o.get('lastName','')}".strip() or o.get("email", f"Owner {oid}")
        nxt = data.get("paging",{}).get("next",{}).get("after")
        if not nxt: break
        params["after"] = nxt
    return out

def get_stage_metadata(token):
    """
    Vrací:
      label_map: stage_id -> label
      order:     pořadí labelů dle první pipeline
      won_ids:   stage_id, kde metadata.isClosed=='true' A probability>=0.5
      lost_ids:  stage_id, kde metadata.isClosed=='true' A probability<0.5

    DŮLEŽITÉ: uzavřenost fáze se pozná podle 'isClosed', NE podle probability.
    Otevřené fáze mohou mít probability 0% (např. 'Not Now'/'Awaiting
    confirmation') nebo blízko 100%, aniž by byly ve skutečnosti uzavřené -
    probability samo o sobě NENÍ spolehlivý signál uzavřenosti.
    """
    data = _get(f"{HUBSPOT_BASE}/crm/v3/pipelines/deals", token)
    label_map, order, won_ids, lost_ids = {}, [], set(), set()
    for pipe in data.get("results", []):
        stages = pipe.get("stages", [])
        if not order: order = [(s.get("label") or s.get("id")).strip() for s in stages]
        for s in stages:
            sid = s.get("id"); label_map[sid] = (s.get("label") or sid).strip()
            meta = s.get("metadata") or {}
            is_closed = str(meta.get("isClosed", "")).strip().lower() == "true"
            if not is_closed:
                continue
            prob = meta.get("probability")
            try:
                pf = float(prob) if prob is not None else 0.0
            except ValueError:
                pf = 0.0
            if pf >= 0.5: won_ids.add(sid)
            else: lost_ids.add(sid)
    return label_map, order, won_ids, lost_ids

def _search_deals(token, filters, extra_props=None):
    url = f"{HUBSPOT_BASE}/crm/v3/objects/deals/search"
    props = ["dealname","dealstage","amount","hubspot_owner_id","closedate","pipeline"]
    if extra_props: props += extra_props
    body = {
        "filterGroups": [{"filters": filters}],
        "properties": props,
        "limit": 100,
        "sorts": [{"propertyName": "hs_lastmodifieddate", "direction": "DESCENDING"}],
    }
    deals = []; after = None
    while True:
        if after: body["after"] = after
        data = _post(url, token, body); deals.extend(data.get("results", []))
        after = data.get("paging",{}).get("next",{}).get("after")
        if not after: break
    return deals

def fetch_deals_by_closedate_cutoff(token, cutoff_epoch_ms):
    """
    Přesná replika obou původních skriptů (hubspot_weekly_report_2026.py
    a hubspot_weekly_report_dynamic_2026.py): stáhne VŠECHNY dealy (bez
    ohledu na stage - otevřené i uzavřené) s closedate < cutoff. ŽÁDNÝ
    filtr na stage, žádná spodní hranice - přesně jak to dělaly oba
    staré skripty.

    Cutoff se volá dvakrát s různou hodnotou:
      - konec kalendářního roku  -> "Pipeline till end of year" základ
      - dnešek + 18 měsíců       -> "Rolling 18" základ (superset toho výše)
    Protože "18 měsíců dopředu" je vždy později než "konec roku", stačí
    zavolat JEDNOU s širším (18měsíčním) cutoffem a užší množinu odvodit
    lokálně filtrem na closedate - přesně to dělá main.py.
    """
    filters = [{"propertyName": "closedate", "operator": "LT", "value": cutoff_epoch_ms}]
    return _search_deals(token, filters)

def fetch_all_open_deals_no_date_filter(token, closed_stage_ids):
    """
    DIAGNOSTICKÁ funkce (jen pro log, nepoužívá se pro report). Stáhne
    VŠECHNY aktuálně otevřené dealy BEZ ohledu na closedate. Porovnáním
    s fetch_deals_by_closedate_cutoff() se pozná, kolik dealů vypadává
    kvůli chybějícímu/vzdálenému closedate.
    """
    filters = []
    if closed_stage_ids:
        filters.append({"propertyName": "dealstage", "operator": "NOT_IN", "values": list(closed_stage_ids)})
    return _search_deals(token, filters) if filters else []

def fetch_closed_deals_since(token, closed_stage_ids, since_epoch_ms):
    """Won/Lost dealy s closedate >= since_epoch_ms. Uzavřené dealy VŽDY mají
    closedate vyplněné (je to datum uzavření), takže filtr je tu spolehlivý."""
    if not closed_stage_ids:
        return []
    filters = [
        {"propertyName": "dealstage", "operator": "IN", "values": list(closed_stage_ids)},
        {"propertyName": "closedate", "operator": "GTE", "value": since_epoch_ms},
    ]
    return _search_deals(token, filters)

def get_deal_company_ids(token, deal_ids):
    if not deal_ids: return {}
    url = f"{HUBSPOT_BASE}/crm/v4/associations/deals/companies/batch/read"
    result = {}
    unique_ids = list(dict.fromkeys(deal_ids))
    for i in range(0, len(unique_ids), 100):
        chunk = unique_ids[i:i+100]
        data = _post(url, token, {"inputs": [{"id": d} for d in chunk]})
        for r in data.get("results", []):
            fid = str((r.get("from") or {}).get("id")); to = r.get("to") or []
            if fid and to: result[fid] = str(to[0].get("toObjectId"))
    return result

def get_company_names(token, company_ids):
    if not company_ids: return {}
    url = f"{HUBSPOT_BASE}/crm/v3/objects/companies/batch/read"
    names = {}
    unique_ids = list(dict.fromkeys(company_ids))
    for i in range(0, len(unique_ids), 100):
        chunk = unique_ids[i:i+100]
        data = _post(url, token, {"properties": ["name"], "inputs": [{"id": c} for c in chunk]})
        for r in data.get("results", []):
            names[str(r.get("id"))] = (r.get("properties") or {}).get("name", "")
    return names

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
    data = _get(f"{HUBSPOT_BASE}/crm/v3/pipelines/deals", token)
    label_map, order, won_ids, lost_ids = {}, [], set(), set()
    for pipe in data.get("results", []):
        stages = pipe.get("stages", [])
        if not order: order = [s.get("label", s.get("id")) for s in stages]
        for s in stages:
            sid = s.get("id"); label_map[sid] = s.get("label", sid)
            p = (s.get("metadata") or {}).get("probability")
            if p is not None:
                try:
                    pf = float(p)
                    if pf >= 0.999: won_ids.add(sid)
                    elif pf <= 0.001: lost_ids.add(sid)
                except ValueError: pass
    return label_map, order, won_ids, lost_ids

def fetch_deals(token, cutoff_epoch_ms):
    url = f"{HUBSPOT_BASE}/crm/v3/objects/deals/search"
    body = {
        "filterGroups": [{"filters": [
                    {"propertyName": "closedate", "operator": "GTE", "value": 1767225600000},  # 2026-01-01 00:00 UTC
                    {"propertyName": "closedate", "operator": "LT",  "value": cutoff_epoch_ms}
                ]}],
        "properties": ["dealname","dealstage","amount","hubspot_owner_id","closedate","pipeline"],
        "limit": 100, "sorts": [{"propertyName": "closedate", "direction": "DESCENDING"}],
    }
    deals = []; after = None
    while True:
        if after: body["after"] = after
        data = _post(url, token, body); deals.extend(data.get("results", []))
        after = data.get("paging",{}).get("next",{}).get("after")
        if not after: break
    return deals

def get_deal_company_ids(token, deal_ids):
    if not deal_ids: return {}
    url = f"{HUBSPOT_BASE}/crm/v4/associations/deals/companies/batch/read"
    result = {}
    for i in range(0, len(deal_ids), 100):
        chunk = list(dict.fromkeys(deal_ids))[i:i+100]
        data = _post(url, token, {"inputs": [{"id": d} for d in chunk]})
        for r in data.get("results", []):
            fid = str((r.get("from") or {}).get("id")); to = r.get("to") or []
            if fid and to: result[fid] = str(to[0].get("toObjectId"))
    return result

def get_company_names(token, company_ids):
    if not company_ids: return {}
    url = f"{HUBSPOT_BASE}/crm/v3/objects/companies/batch/read"
    names = {}
    for i in range(0, len(company_ids), 100):
        chunk = list(dict.fromkeys(company_ids))[i:i+100]
        data = _post(url, token, {"properties": ["name"], "inputs": [{"id": c} for c in chunk]})
        for r in data.get("results", []):
            names[str(r.get("id"))] = (r.get("properties") or {}).get("name", "")
    return names

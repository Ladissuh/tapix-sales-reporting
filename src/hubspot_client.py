"""
Veškerá komunikace s HubSpot API na jednom místě.
Nahrazuje logiku z obou dnešních skriptů (hubspot_weekly_report_2026.py
a hubspot_weekly_report_dynamic_2026.py) - stahuje se ale jen JEDNOU,
s nejširším (18měsíčním) cutoffem; užší "do konce roku" pohled se filtruje
až lokálně v metrics.py.
"""
import os
import time
from typing import Dict, List, Tuple

import requests

HUBSPOT_BASE = "https://api.hubapi.com"


def load_token() -> str:
    token = os.getenv("HUBSPOT_TOKEN")
    if not token:
        raise RuntimeError(
            "Chybí HUBSPOT_TOKEN. Lokálně: založ .env s HUBSPOT_TOKEN=pat-..., "
            "na GitHubu: Settings -> Secrets and variables -> Actions."
        )
    return token


def _headers(token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _backoff_sleep(attempt: int) -> None:
    time.sleep(min(2 ** attempt, 32))


def _get_with_retry(url: str, token: str, params: dict = None) -> dict:
    attempt = 0
    while True:
        resp = requests.get(url, headers=_headers(token), params=params)
        if resp.status_code == 429 or 500 <= resp.status_code < 600:
            attempt += 1
            _backoff_sleep(attempt)
            continue
        resp.raise_for_status()
        return resp.json()


def _post_with_retry(url: str, token: str, body: dict) -> dict:
    attempt = 0
    while True:
        resp = requests.post(url, headers=_headers(token), json=body)
        if resp.status_code == 429 or 500 <= resp.status_code < 600:
            attempt += 1
            _backoff_sleep(attempt)
            continue
        resp.raise_for_status()
        return resp.json()


def get_all_owners(token: str) -> Dict[str, str]:
    """owner_id -> display name"""
    url = f"{HUBSPOT_BASE}/crm/v3/owners/"
    owners_map: Dict[str, str] = {}
    params = {"limit": 100, "archived": "false"}
    while True:
        data = _get_with_retry(url, token, params)
        for o in data.get("results", []):
            owner_id = str(o.get("id"))
            name = (
                f"{o.get('firstName', '')} {o.get('lastName', '')}".strip()
                or o.get("email", f"Owner {owner_id}")
            )
            owners_map[owner_id] = name
        next_after = data.get("paging", {}).get("next", {}).get("after")
        if not next_after:
            break
        params["after"] = next_after
    return owners_map


def get_stage_metadata(token: str) -> Tuple[Dict[str, str], List[str], set, set]:
    """
    Vrací:
      stage_label_map: stage_id -> label
      default_order:   pořadí labelů dle první pipeline
      closed_won_ids:  set stage_id, které HubSpot eviduje jako Closed Won
      closed_lost_ids: set stage_id, které HubSpot eviduje jako Closed Lost

    Detekce Won/Lost jde přes stage metadata "probability" (HubSpot
    konvence: 1.0 = closed won, 0.0 = closed lost). Pokud váš pipeline
    tuhle konvenci nedodržuje, uprav podmínku níž.
    """
    url = f"{HUBSPOT_BASE}/crm/v3/pipelines/deals"
    data = _get_with_retry(url, token)

    stage_label_map: Dict[str, str] = {}
    default_order: List[str] = []
    closed_won_ids, closed_lost_ids = set(), set()

    for pipe in data.get("results", []):
        stages = pipe.get("stages", [])
        if not default_order:
            default_order = [s.get("label", s.get("id")) for s in stages]
        for s in stages:
            sid = s.get("id")
            stage_label_map[sid] = s.get("label", sid)
            prob = (s.get("metadata") or {}).get("probability")
            if prob is not None:
                try:
                    p = float(prob)
                    if p >= 0.999:
                        closed_won_ids.add(sid)
                    elif p <= 0.001:
                        closed_lost_ids.add(sid)
                except ValueError:
                    pass

    return stage_label_map, default_order, closed_won_ids, closed_lost_ids


def fetch_deals(token: str, cutoff_epoch_ms: int) -> List[dict]:
    """
    Stáhne VŠECHNY dealy s closedate < cutoff, včetně dealname a asociované
    company (přes 'associations' parametr - žádné extra volání navíc).
    """
    url = f"{HUBSPOT_BASE}/crm/v3/objects/deals/search"
    body = {
        "filterGroups": [
            {
                "filters": [
                    {"propertyName": "closedate", "operator": "LT", "value": cutoff_epoch_ms}
                ]
            }
        ],
        "properties": ["dealname", "dealstage", "amount", "hubspot_owner_id", "closedate", "pipeline"],
        "associations": ["companies"],
        "limit": 100,
        "sorts": [{"propertyName": "closedate", "direction": "DESCENDING"}],
    }

    all_deals: List[dict] = []
    after = None
    while True:
        if after:
            body["after"] = after
        data = _post_with_retry(url, token, body)
        all_deals.extend(data.get("results", []))
        after = data.get("paging", {}).get("next", {}).get("after")
        if not after:
            break
    return all_deals


def get_company_names(token: str, company_ids: List[str]) -> Dict[str, str]:
    """Batch-read názvů firem pro dané company_id (max 100 na dávku)."""
    if not company_ids:
        return {}
    url = f"{HUBSPOT_BASE}/crm/v3/objects/companies/batch/read"
    names: Dict[str, str] = {}
    unique_ids = list(dict.fromkeys(company_ids))
    for i in range(0, len(unique_ids), 100):
        chunk = unique_ids[i:i + 100]
        body = {"properties": ["name"], "inputs": [{"id": cid} for cid in chunk]}
        data = _post_with_retry(url, token, body)
        for r in data.get("results", []):
            names[str(r.get("id"))] = (r.get("properties") or {}).get("name", "")
    return names


def extract_primary_company_id(deal: dict) -> str:
    assoc = (deal.get("associations") or {}).get("companies") or {}
    results = assoc.get("results") or []
    if results:
        return str(results[0].get("id"))
    return ""

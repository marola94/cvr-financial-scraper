"""
Part 2 — CVR API lookup.

For each input row, determines the best search strategy and returns ALL
matching companies as a list.

Search priority:
  CVR            → exact match, returns 1 company
  Company        → company name search, returns all matches
  Contact person → person name search, returns all matches
                   CPR (birth date) and/or role narrow the search if provided
                   CPR alone is not a valid standalone search
"""
import re
import time
import requests
import config


# ── Elasticsearch helpers ──────────────────────────────────────────────────────

def _post(query: dict) -> list:
    """Send a query to the CVR Elasticsearch API and return the hits list."""
    try:
        resp = requests.post(
            config.CVR_SEARCH_URL,
            json=query,
            auth=config.CVR_AUTH,
            timeout=30,
        )
        resp.raise_for_status()
    except requests.RequestException as e:
        print(f"  [CVR] API-fejl: {e}")
        return []
    return resp.json().get("hits", {}).get("hits", [])


# ── Search functions ───────────────────────────────────────────────────────────

def search_by_cvr(cvr: str) -> list:
    return _post({
        "size": 1,
        "query": {"term": {"Vrvirksomhed.cvrNummer": int(cvr)}},
    })


def search_by_company_name(name: str) -> list:
    return _post({
        "size": 3000,
        "query": {
            "match": {
                "Vrvirksomhed.virksomhedMetadata.nyesteNavn.navn": {
                    "query": name,
                    "operator": "and",
                }
            }
        },
    })


def search_by_person(name: str, role: str = None) -> list:
    """
    Search by contact person name, optionally narrowed by role.
    """
    must = [
        {"match": {"Vrvirksomhed.deltagerRelation.deltager.navne.navn": name}}
    ]
    if role:
        must.append({
            "match": {
                "Vrvirksomhed.deltagerRelation.organisationer.organisationsNavn": role
            }
        })
    return _post({
        "size": 3000,
        "query": {"bool": {"must": must}},
    })


# ── Data extraction ────────────────────────────────────────────────────────────

def _safe(d, *keys, default=""):
    """Safe nested dict access — returns default if any key is missing."""
    for k in keys:
        if not isinstance(d, dict):
            return default
        d = d.get(k)
        if d is None:
            return default
    return d if d is not None else default


def _region_from_postnr(postnr) -> str:
    try:
        p = int(str(postnr).strip())
    except (ValueError, TypeError):
        return ""
    if 1000 <= p <= 2999 or 3000 <= p <= 3699 or 3770 <= p <= 3799:
        return "Region Hovedstaden"
    if 3700 <= p <= 3769 or 4000 <= p <= 4999:
        return "Region Sjælland"
    if 5000 <= p <= 6999:
        return "Region Syddanmark"
    if 7000 <= p <= 8999:
        return "Region Midtjylland"
    if 9000 <= p <= 9999:
        return "Region Nordjylland"
    return ""


def _active_name(navne: list) -> str:
    """Return the currently active name from a CVR name-history list."""
    active = [n for n in (navne or [])
              if isinstance(n, dict) and n.get("periode", {}).get("gyldigTil") is None]
    source = active if active else (navne or [])
    return source[-1].get("navn", "") if source else ""


def _extract_persons(relations: list) -> tuple:
    """
    Parse deltagerRelation and return three lists:
    (board_members, directors, owners) — active records only.
    """
    bestyrelse, direktion, ejere = [], [], []

    for rel in (relations or []):
        if not isinstance(rel, dict):
            continue
        if rel.get("periode", {}).get("gyldigTil") is not None:
            continue

        deltager = rel.get("deltager") or {}
        navn = _active_name(deltager.get("navne", []))
        if not navn:
            continue

        for org in (rel.get("organisationer") or []):
            if not isinstance(org, dict):
                continue
            if org.get("periode", {}).get("gyldigTil") is not None:
                continue

            org_navn_raw = org.get("organisationsNavn")
            if isinstance(org_navn_raw, list):
                active = [n for n in org_navn_raw
                          if isinstance(n, dict)
                          and n.get("periode", {}).get("gyldigTil") is None]
                source = active if active else org_navn_raw
                rolle  = source[-1].get("navn", "").lower() if source else ""
            elif isinstance(org_navn_raw, str):
                rolle = org_navn_raw.lower()
            else:
                rolle = ""

            if "bestyrelse" in rolle:
                if navn not in bestyrelse:
                    bestyrelse.append(navn)
            elif "direktion" in rolle or "direktør" in rolle:
                if navn not in direktion:
                    direktion.append(navn)
            elif "ejer" in rolle or "legale ejere" in rolle:
                if navn not in ejere:
                    ejere.append(navn)

    return bestyrelse, direktion, ejere


def extract_company_data(hit: dict) -> dict:
    """Parse a raw CVR API hit into a normalised company dict."""
    src  = hit.get("_source", {}).get("Vrvirksomhed", {})
    meta = src.get("virksomhedMetadata", {}) or {}

    navn_obj = meta.get("nyesteNavn") or {}
    navn = navn_obj.get("navn", "") if isinstance(navn_obj, dict) else ""

    adr     = meta.get("nyesteBeliggenhedsadresse") or {}
    postnr  = adr.get("postnummer", "")
    by      = adr.get("postDistrikt", "")
    kommune = adr.get("kommuneNavn", "")
    region  = _region_from_postnr(postnr)

    branche_obj = meta.get("nyesteHovedbranche") or {}
    branche     = branche_obj.get("branchetekst", "")
    branchekode = branche_obj.get("branchekode", "")

    form_obj  = meta.get("nyesteVirksomhedsform") or {}
    virk_form = form_obj.get("kortBeskrivelse", "")

    status = meta.get("sammensatStatus", "")

    telefon = email = hjemmeside = ""
    for k in (meta.get("nyesteKontaktoplysninger") or []):
        if not isinstance(k, dict):
            continue
        ktype = k.get("kontaktoplysningstype", "")
        kval  = k.get("kontaktoplysning", "")
        if ktype == "TELEFONNUMMER" and not telefon:
            telefon = kval
        elif ktype == "EMAILADRESSE" and not email:
            email = kval
        elif ktype == "HJEMMESIDE" and not hjemmeside:
            hjemmeside = kval

    ansatte_obj = meta.get("nyesteAarsbeskaeftigelse") or {}
    ansatte     = ansatte_obj.get("antalAnsatteInterval", "")
    ansatte_aar = ansatte_obj.get("aar", "")

    stiftelse = meta.get("stiftelsesDato") or ""
    if not stiftelse:
        livsforloeb = src.get("livsforloeb") or []
        if livsforloeb:
            stiftelse = _safe(livsforloeb[0], "periode", "gyldigFra")

    bestyrelse, direktion, ejere = _extract_persons(
        src.get("deltagerRelation", [])
    )

    return {
        "cvr":             src.get("cvrNummer", ""),
        "name":            navn,
        "city":            by,
        "postal_code":     str(postnr),
        "municipality":    kommune,
        "region":          region,
        "industry":        branche,
        "industry_code":   branchekode,
        "virksomhedsform": virk_form,
        "status":          status,
        "phone":           telefon,
        "email":           email,
        "website":         hjemmeside,
        "employees":       ansatte,
        "employees_year":  str(ansatte_aar),
        "founded":         str(stiftelse),
        "board_members":   bestyrelse,
        "directors":       direktion,
        "owners":          ejere,
        "_raw":            src,
    }


# ── Main lookup ────────────────────────────────────────────────────────────────

def lookup_companies(row: dict) -> list:
    """
    Determine search strategy from available row data.
    Returns a list of company dicts — one per match found.
    """
    cvr          = re.sub(r"\D", "", str(row.get("CVR") or ""))
    company_name = (row.get("Company") or "").strip()
    contact      = (row.get("Contact person") or "").strip()
    contact_role = (row.get(config.ROLE_COLUMN) or "").strip()

    hits      = []
    soegt_via = None

    if cvr and len(cvr) == 8 and cvr.isdigit():
        print(f"  [CVR] Henter virksomhedsdata via CVR {cvr}…")
        hits      = search_by_cvr(cvr)
        soegt_via = "CVR"

    elif company_name:
        print(f"  [CVR] Søger virksomheder med navn '{company_name}'…")
        hits      = search_by_company_name(company_name)
        soegt_via = "Company"

    elif contact:
        filters = [f"navn='{contact}'"]
        if contact_role:
            filters.append(f"rolle='{contact_role}'")
        print(f"  [CVR] Søger kontaktperson ({', '.join(filters)})…")
        hits      = search_by_person(contact, role=contact_role or None)
        soegt_via = "Contact person"

    time.sleep(config.SLEEP_BETWEEN_CALLS)

    if not hits:
        print(f"  [CVR] Ingen match fundet")
        return [_empty_company(cvr, company_name or contact, soegt_via)]

    print(f"  [CVR] {len(hits)} virksomhed(er) fundet")
    total = len(hits)
    return [
        {**extract_company_data(hit), "_soegt_via": soegt_via, "_antal_matches": total}
        for hit in hits
    ]


def _empty_company(cvr: str, name: str, soegt_via) -> dict:
    fields = [
        "cvr", "name", "city", "postal_code", "municipality", "region",
        "industry", "industry_code", "virksomhedsform", "status",
        "phone", "email", "website", "employees", "employees_year",
        "founded",
    ]
    d = {f: "" for f in fields}
    d.update({
        "cvr": cvr, "name": name,
        "board_members": [], "directors": [], "owners": [],
        "_raw": {}, "_soegt_via": soegt_via, "_antal_matches": 0,
    })
    return d

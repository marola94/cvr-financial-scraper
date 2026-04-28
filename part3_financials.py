"""
Part 3 — Fetch financial reports for a company.

Queries the offentliggoerelser Elasticsearch index for XBRL documents,
downloads each document, and parses the financial figures.
Returns up to YEARS_OF_FINANCIALS annual reports, newest first.

Field resolution uses taxonomy_mappings.json — FSA checked first, then IFRS.
Each element is matched on (namespace, name) to avoid cross-taxonomy confusion.
"""
import gzip
import json
import os
import xml.etree.ElementTree as ET
from datetime import date
import requests
import time
import config


# ── Namespaces ─────────────────────────────────────────────────────────────────
XBRLI_NS    = "http://www.xbrl.org/2003/instance"
LINKBASE_NS = "http://www.xbrl.org/2003/linkbase"
XLINK_NS    = "http://www.w3.org/1999/xlink"

METADATA_NS = {
    "http://xbrl.dcca.dk/cmn",
    "http://xbrl.dcca.dk/gsd",
    "http://xbrl.dcca.dk/sob",
    "http://xbrl.dcca.dk/arr",
    "http://xbrl.dcca.dk/mrv",
    XBRLI_NS,
    LINKBASE_NS,
    "http://www.w3.org/1999/xlink",
    "http://xbrl.org/2006/xbrldi",
    "",
}

# ── Load taxonomy mappings once at module level ────────────────────────────────
def _base_dir() -> str:
    """Return the directory containing bundled data files (works in PyInstaller exe and dev)."""
    import sys
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(__file__)

_MAPPINGS_PATH = os.path.join(_base_dir(), "taxonomy_mappings.json")
with open(_MAPPINGS_PATH, encoding="utf-8") as _f:
    _MAPPINGS = json.load(_f)

_NS_PATTERNS  = _MAPPINGS["namespace_patterns"]   # {"fsa": "http://...", "ifrs": "ifrs"}
_TEXT_FIELDS  = _MAPPINGS["text_fields"]           # {"fsa": {"ClassOfReportingEntity": "reporting_class"}}
_FIELD_DEFS   = _MAPPINGS["fields"]                # all field definitions
_TAX_PRIORITY = list(_NS_PATTERNS.keys())          # ["fsa", "ifrs"] — order matters


def _resolve_ns(ns_uri: str) -> str:
    """Map a full namespace URI to its short prefix, or 'other' if unknown."""
    for prefix, pattern in _NS_PATTERNS.items():
        if pattern in ns_uri:
            return prefix
    return "other"


# ── Public function ────────────────────────────────────────────────────────────

def fetch_financials(cvr: str) -> tuple:
    """Returns (reports: list, filing_type: str)."""
    # Query all filings without mime-type filter to detect PDF vs XBRL
    query_all = {
        "query": {"bool": {"must": [{"term": {"cvrNummer": int(cvr)}}]}},
        "size": 5,
        "sort": [{"offentliggoerelsesTidspunkt": {"order": "desc"}}],
    }
    # Query filtered to XML only for parsing
    query_xml = {
        "query": {
            "bool": {
                "must": [
                    {"term": {"cvrNummer": int(cvr)}},
                    {"term": {"dokumenter.dokumentMimeType": "application"}},
                    {"term": {"dokumenter.dokumentMimeType": "xml"}},
                ]
            }
        },
        "size": 20,
        "sort": [{"offentliggoerelsesTidspunkt": {"order": "desc"}}],
    }

    try:
        resp_all = requests.post(config.REGNSKAB_SEARCH_URL, json=query_all, timeout=30)
        resp_all.raise_for_status()
        resp_xml = requests.post(config.REGNSKAB_SEARCH_URL, json=query_xml, timeout=30)
        resp_xml.raise_for_status()
    except requests.RequestException as e:
        print(f"  [Regnskab] Søgefejl: {e}")
        return [], "Fejl"

    all_hits = resp_all.json().get("hits", {}).get("hits", [])
    hits     = resp_xml.json().get("hits", {}).get("hits", [])

    # Determine filing type from all filings
    if not all_hits:
        filing_type = "Ingen regnskaber"
    elif hits:
        filing_type = "XBRL"
    else:
        mime_types = set()
        for hit in all_hits:
            for dok in hit.get("_source", {}).get("dokumenter", []):
                if isinstance(dok, dict):
                    mime_types.add(dok.get("dokumentMimeType", "").lower())
        if any("pdf" in m for m in mime_types):
            filing_type = "PDF"
        else:
            filing_type = "Ukendt format"

    reports    = []
    seen_years = set()

    for hit in hits:
        if len(reports) >= config.YEARS_OF_FINANCIALS:
            break

        xbrl_url = _find_xbrl_url(hit.get("_source", {}).get("dokumenter", []))
        if not xbrl_url:
            continue

        try:
            xbrl_resp = requests.get(xbrl_url, timeout=30)
            xbrl_resp.raise_for_status()
            parsed = _parse_xbrl(xbrl_resp.content)
        except requests.RequestException as e:
            print(f"  [XBRL] Download fejl: {e}")
            continue

        period_end = parsed.get("period_end", "")
        if not period_end:
            continue

        year = period_end[:4]
        if year in seen_years:
            continue
        seen_years.add(year)

        ebit   = parsed.get("ebit")
        da_val = parsed.get("depreciation")
        ebitda = (ebit + da_val) if (ebit is not None and da_val is not None) else None

        reports.append({
            "year":                       year,
            "period_end":                 period_end,
            "currency":                   parsed.get("currency", "DKK"),
            "taxonomy_version":           parsed.get("taxonomy_version", ""),
            "reporting_class":            parsed.get("reporting_class", ""),
            # Income statement
            "revenue":                    parsed.get("revenue"),
            "gross_profit":               parsed.get("gross_profit"),
            "ebit":                       ebit,
            "ebit_elements":              parsed.get("ebit_elements", ""),
            "ebit_is_fsa":                parsed.get("ebit_is_fsa"),
            "depreciation":               da_val,
            "depreciation_elements":      parsed.get("depreciation_elements", ""),
            "depreciation_is_fsa":        parsed.get("depreciation_is_fsa"),
            "depreciation_element_count": parsed.get("depreciation_element_count"),
            "ebitda":                     ebitda,
            "net_profit":                 parsed.get("net_profit"),
            "employees_xbrl":             parsed.get("employees_xbrl"),
            "cost_of_sales":              parsed.get("cost_of_sales"),
            "other_operating_income":     parsed.get("other_operating_income"),
            "external_expenses":          parsed.get("external_expenses"),
            "employee_benefits":          parsed.get("employee_benefits"),
            "wages_salaries":             parsed.get("wages_salaries"),
            "tax_expense":                parsed.get("tax_expense"),
            "current_tax":                parsed.get("current_tax"),
            "deferred_tax":               parsed.get("deferred_tax"),
            "finance_income":             parsed.get("finance_income"),
            "finance_expenses":           parsed.get("finance_expenses"),
            "profit_before_tax":          parsed.get("profit_before_tax"),
            # Balance sheet
            "assets":                     parsed.get("assets"),
            "equity":                     parsed.get("equity"),
            "current_assets":             parsed.get("current_assets"),
            "noncurrent_assets":          parsed.get("noncurrent_assets"),
            "cash":                       parsed.get("cash"),
            "shortterm_receivables":      parsed.get("shortterm_receivables"),
            "shortterm_liabilities":      parsed.get("shortterm_liabilities"),
            "longterm_liabilities":       parsed.get("longterm_liabilities"),
            "provisions":                 parsed.get("provisions"),
            "contributed_capital":        parsed.get("contributed_capital"),
            "retained_earnings":          parsed.get("retained_earnings"),
            "proposed_dividend":          parsed.get("proposed_dividend"),
            # Raw data for sheets 4 & 5
            "_xbrl_raw":                  parsed.get("_xbrl_raw", {}),
            "_financial_raw":             parsed.get("_financial_raw", {}),
            "_financial_ns":              parsed.get("_financial_ns", {}),
            "_misc_raw":                  parsed.get("_misc_raw", {}),
            "_main_context_id":           parsed.get("_main_context_id", ""),
        })

        time.sleep(config.SLEEP_BETWEEN_CALLS)

    reports.sort(key=lambda x: x["period_end"], reverse=True)
    return reports, filing_type


# ── Helpers ────────────────────────────────────────────────────────────────────

def _find_xbrl_url(dokumenter: list):
    for dok in (dokumenter or []):
        if not isinstance(dok, dict):
            continue
        mime = dok.get("dokumentMimeType", "")
        if "xml" in mime.lower():
            return dok.get("dokumentUrl")
    return None


def _parse_xbrl(content: bytes) -> dict:
    """
    Parse XBRL XML bytes and extract all financial figures.

    Two-pass approach:
      Pass 1 — build element_map: {(ns_prefix, local_name): value} for all
               main-context elements, with namespace resolved per element.
      Pass 2 — resolve mapped fields from taxonomy_mappings.json, trying
               each taxonomy in priority order (fsa → ifrs).
    """
    if not content.lstrip().startswith(b"<"):
        try:
            content = gzip.decompress(content)
        except Exception:
            pass

    try:
        root = ET.fromstring(content)
    except ET.ParseError as e:
        print(f"  [XBRL] Parse-fejl: {e}")
        return {}

    result = {}

    # ── Taxonomy version from schemaRef ───────────────────────────────────────
    schema = root.find(f"{{{LINKBASE_NS}}}schemaRef")
    if schema is not None:
        result["taxonomy_version"] = schema.get(f"{{{XLINK_NS}}}href", "")

    # ── Currency from first monetary unit ─────────────────────────────────────
    for unit in root.iter(f"{{{XBRLI_NS}}}unit"):
        measure = unit.findtext(f"{{{XBRLI_NS}}}measure", "")
        if "iso4217:" in measure:
            result["currency"] = measure.split("iso4217:")[-1].strip()
            break

    # ── Context map — skip scenario contexts ──────────────────────────────────
    contexts = {}
    for ctx in root.iter(f"{{{XBRLI_NS}}}context"):
        ctx_id = ctx.get("id", "")
        if ctx.find(f"{{{XBRLI_NS}}}scenario") is not None:
            continue
        period = ctx.find(f"{{{XBRLI_NS}}}period")
        if period is None:
            continue
        start = period.findtext(f"{{{XBRLI_NS}}}startDate")
        end   = period.findtext(f"{{{XBRLI_NS}}}endDate")
        if start and end:
            contexts[ctx_id] = (start, end)

    if not contexts:
        return result

    def _days(pair):
        try:
            return (date.fromisoformat(pair[1]) - date.fromisoformat(pair[0])).days
        except Exception:
            return 0

    main_id       = max(contexts, key=lambda k: _days(contexts[k]))
    _, period_end = contexts[main_id]
    result.update({"period_end": period_end, "_main_context_id": main_id})

    # ── Pass 1: collect all elements ──────────────────────────────────────────
    element_map   = {}   # (ns_prefix, local) -> value  — used for field resolution
    xbrl_raw      = {}   # local -> value               — all elements (for _xbrl_raw)
    financial_raw = {}   # local -> value               — non-metadata numeric elements
    financial_ns  = {}   # local -> ns_uri
    misc_raw      = {}   # local -> value               — text or metadata elements

    for elem in root.iter():
        if "}" in elem.tag:
            ns_uri, local = elem.tag[1:].split("}", 1)
        else:
            ns_uri, local = "", elem.tag

        ns_prefix = _resolve_ns(ns_uri)

        # Text fields — extracted without strict context requirement
        if ns_prefix in _TEXT_FIELDS and local in _TEXT_FIELDS[ns_prefix]:
            if elem.text and elem.text.strip():
                result[_TEXT_FIELDS[ns_prefix][local]] = elem.text.strip()
            continue

        if elem.get("contextRef") != main_id:
            continue

        text = (elem.text or "").strip()
        if not text:
            continue

        try:
            val = float(text)
            is_numeric = True
        except ValueError:
            val = text
            is_numeric = False

        element_map[(ns_prefix, local)] = val
        xbrl_raw[local] = val

        if is_numeric and ns_uri not in METADATA_NS:
            financial_raw[local] = val
            financial_ns[local]  = ns_uri
        else:
            misc_raw[local] = val

    # ── Pass 2: resolve mapped fields from JSON ───────────────────────────────
    for field, field_def in _FIELD_DEFS.items():
        track = field_def.get("track_source", False)

        for taxonomy in _TAX_PRIORITY:
            if taxonomy not in field_def:
                continue

            found = False
            for group in field_def[taxonomy]["groups"]:
                method   = group["method"]
                elements = group["elements"]

                if method == "first":
                    for elem_name in elements:
                        val = element_map.get((taxonomy, elem_name))
                        if val is not None:
                            result[field] = val
                            if track:
                                result[f"{field}_elements"]      = elem_name
                                result[f"{field}_is_fsa"]        = (taxonomy == "fsa")
                                result[f"{field}_element_count"] = 1
                            found = True
                            break

                elif method == "sum":
                    found_parts = {
                        e: element_map[(taxonomy, e)]
                        for e in elements
                        if (taxonomy, e) in element_map
                    }
                    if found_parts:
                        result[field] = sum(found_parts.values())
                        if track:
                            result[f"{field}_elements"]      = ", ".join(sorted(found_parts.keys()))
                            result[f"{field}_is_fsa"]        = (taxonomy == "fsa")
                            result[f"{field}_element_count"] = len(found_parts)
                        found = True

                if found:
                    break

            if found:
                break

    result["_xbrl_raw"]      = xbrl_raw
    result["_financial_raw"] = financial_raw
    result["_financial_ns"]  = financial_ns
    result["_misc_raw"]      = misc_raw
    return result

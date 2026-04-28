"""
Part 4 — Calculate financial KPIs from annual reports.

Input:  list of report dicts, newest first (from part3_financials)
Output: dict of calculated KPIs
"""


def _linreg_trend(reports: list, field: str) -> tuple:
    """
    Fit a line through all available (year, value) pairs for the given field.
    Returns (slope > 0: bool, years_str: str) or (None, None) if fewer than 2 points.
    Works regardless of sign — no complex number risk.
    """
    pairs = sorted(
        [(int(r["year"]), r[field]) for r in reports
         if r.get(field) is not None and r.get("year")],
        key=lambda x: x[0],
    )
    if len(pairs) < 2:
        return None, None
    xs = [p[0] for p in pairs]
    ys = [p[1] for p in pairs]
    n = len(pairs)
    x_mean = sum(xs) / n
    y_mean = sum(ys) / n
    num = sum((x - x_mean) * (y - y_mean) for x, y in zip(xs, ys))
    den = sum((x - x_mean) ** 2 for x in xs)
    if den == 0:
        return None, None
    return (num / den) > 0, ", ".join(str(y) for y in xs)


def _cagr(reports: list, field: str) -> tuple:
    """
    CAGR = [(Ending / Beginning)^(1/n) - 1] * 100
    Returns (value, span_str) where value is float or "N/A"
    and span_str is e.g. "2021, 2024 (n=3)".
    Requires: field present in latest report AND at least one older report with same field.
    """
    NA = ("N/A", "N/A")
    pairs = [
        (int(r["year"]), r[field])
        for r in reports
        if r.get(field) is not None and r.get("year")
    ]
    if not pairs:
        return NA
    if pairs[0][0] != int(reports[0]["year"]):
        return NA   # latest report missing the field
    if len(pairs) < 2:
        return NA
    y_new, v_new = pairs[0]
    y_old, v_old = pairs[-1]
    n = y_new - y_old
    if n <= 0 or v_old == 0:
        return NA
    ratio = v_new / v_old
    if ratio < 0:
        return NA
    value = (ratio ** (1 / n) - 1) * 100
    return value, f"{y_old}, {y_new} (n={n})"


def calculate_kpis(reports: list) -> dict:
    kpi = {
        "revenue_cagr":            None,
        "revenue_cagr_span":       None,
        "ebitda_cagr":             None,
        "ebitda_cagr_span":        None,
        "ebit_cagr":               None,
        "ebit_cagr_span":          None,
        "ebitda_margin":           None,
        "rule_of_40":              None,
        "positive_revenue_trend":  None,
        "positive_ebitda_trend":   None,
        "ebitda_trend_period":     None,
        "positive_ebit_trend":     None,
        "ebit_trend_period":       None,
    }

    if not reports:
        return kpi

    latest = reports[0]

    # ── Revenue CAGR ─────────────────────────────────────────────────────────
    kpi["revenue_cagr"], kpi["revenue_cagr_span"] = _cagr(reports, "revenue")

    # ── EBITDA margin (latest) ────────────────────────────────────────────────
    ebitda_latest  = latest.get("ebitda")
    revenue_latest = latest.get("revenue")
    if ebitda_latest is not None and revenue_latest and revenue_latest != 0:
        kpi["ebitda_margin"] = (ebitda_latest / revenue_latest) * 100

    # ── Rule of 40 ────────────────────────────────────────────────────────────
    if isinstance(kpi["revenue_cagr"], float) and kpi["ebitda_margin"] is not None:
        kpi["rule_of_40"] = kpi["revenue_cagr"] + kpi["ebitda_margin"]

    # ── EBITDA CAGR ───────────────────────────────────────────────────────────
    kpi["ebitda_cagr"], kpi["ebitda_cagr_span"] = _cagr(reports, "ebitda")

    # ── EBIT CAGR ─────────────────────────────────────────────────────────────
    kpi["ebit_cagr"], kpi["ebit_cagr_span"] = _cagr(reports, "ebit")

    # ── Positive revenue trend (latest > oldest available) ────────────────────
    revenues = [r.get("revenue") for r in reports if r.get("revenue") is not None]
    if len(revenues) >= 2:
        kpi["positive_revenue_trend"] = revenues[0] > revenues[-1]

    # ── EBITDA trend (linear regression) ─────────────────────────────────────
    kpi["positive_ebitda_trend"], kpi["ebitda_trend_period"] = _linreg_trend(reports, "ebitda")

    # ── EBIT trend (linear regression) ───────────────────────────────────────
    kpi["positive_ebit_trend"], kpi["ebit_trend_period"] = _linreg_trend(reports, "ebit")

    return kpi

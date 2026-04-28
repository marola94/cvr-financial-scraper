"""
Part 1 — Load input file (Excel or CSV).

Accepts .xlsx, .xls, and .csv files.
Requires at least one of: CVR, Company, Contact person, CPR.
Returns a list of row dicts and a list of extra (passthrough) column names.
"""
from pathlib import Path
import pandas as pd
import config

_KEY_COLS = config.SEARCH_PRIORITY + [config.ROLE_COLUMN]


def load_input(filepath: str) -> tuple:
    """
    Load Excel or CSV file.
    Returns (rows: list[dict], extra_cols: list[str]).
    Raises ValueError if no recognised key columns are found.
    """
    path   = Path(filepath)
    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xls"):
        df = pd.read_excel(filepath, dtype=str)
    elif suffix == ".csv":
        df = _load_csv(filepath)
    else:
        raise ValueError(
            f"Ikke-understøttet filtype: '{suffix}'. Brug .xlsx, .xls eller .csv"
        )

    df.columns = [c.strip() for c in df.columns]
    df = df.where(pd.notna(df), None)

    found_keys = [c for c in config.SEARCH_PRIORITY if c in df.columns]
    if not found_keys:
        raise ValueError(
            f"Filen skal have mindst én kolonne fra: {config.SEARCH_PRIORITY}\n"
            f"Fundne kolonner: {list(df.columns)}"
        )

    extra_cols = [c for c in df.columns if c not in _KEY_COLS]

    rows = df.to_dict(orient="records")

    print(f"Indlæst {len(rows)} rækker fra '{path.name}'")
    print(f"  Søgekolonner fundet : {found_keys}")
    if extra_cols:
        print(f"  Ekstra kolonner     : {extra_cols}")

    return rows, extra_cols, df


def _load_csv(filepath: str) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(filepath, dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Kunne ikke bestemme encoding for {filepath}")

"""
CVR Lookup Tool
===============
Reads an Excel/CSV file with company identifiers and enriches each row with
CVR master data and financial KPIs from Erhvervsstyrelsen's public APIs.

Usage:
    python main.py <input_file> [output_file]

Example:
    python main.py firmaliste.xlsx rapport.xlsx

The input file must contain at least one of:
    CVR, Company, Contact person, CPR
"""
import sys
import time

import part1_loader       as loader
import part2_cvr          as cvr_api
import part3_financials   as financials
import part4_calculations as calculations
import part5_reporter     as reporter
import config


def main(input_file: str, output_file: str = "output.xlsx"):
    if not config.CVR_USER or not config.CVR_PASSWORD:
        import os
        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(sys.executable)
        else:
            exe_dir = os.path.dirname(os.path.abspath(__file__))
        env_path = os.path.join(exe_dir, ".env")
        print("FEJL: CVR-credentials mangler.")
        print()
        print("Placer en .env fil i samme mappe som programmet:")
        print(f"  {env_path}")
        print()
        print("Indholdet af .env skal være:")
        print("  CVR_USER=dit_brugernavn")
        print("  CVR_PASSWORD=dit_password")
        print()
        print("Genstart programmet efter du har oprettet filen.")
        sys.exit(1)

    print("=" * 60)
    print("CVR Lookup Tool")
    print("=" * 60)
    print(f"Input:  {input_file}")
    print(f"Output: {output_file}\n")

    start = time.time()

    # ── Part 1: Load input ────────────────────────────────────────────────────
    rows, extra_cols, input_df = loader.load_input(input_file)

    # ── Parts 2–4: Process each row ───────────────────────────────────────────
    companies = []

    for i, row in enumerate(rows):
        label = row.get("Company") or row.get("CVR") or f"Række {i + 1}"
        print(f"\n[{i + 1}/{len(rows)}] {label}")

        # Part 2: CVR lookup — returns list (multiple matches possible)
        matches = cvr_api.lookup_companies(row)

        for company in matches:
            company["_input"] = row

            cvr = str(company.get("cvr", "")).strip()
            if not cvr:
                print("  Ingen CVR fundet — springer regnskaber over")
                company["reports"]     = []
                company["filing_type"] = "Ingen match"
                company["kpi"]         = {}
                companies.append(company)
                continue

            # Part 3: Financial reports
            print(f"  Henter regnskaber for CVR {cvr}…")
            regs, filing_type = financials.fetch_financials(cvr)
            company["reports"]     = regs
            company["filing_type"] = filing_type

            if regs:
                years = [r["year"] for r in regs]
                print(f"  Fandt {len(regs)} regnskabsår: {years}")
            else:
                print("  Ingen XBRL-regnskaber fundet")

            # Part 4: KPIs
            company["kpi"] = calculations.calculate_kpis(regs)

            companies.append(company)

    # ── Part 5: Excel output ──────────────────────────────────────────────────
    print(f"\nSkriver Excel-fil med {len(companies)} virksomheder…")
    reporter.create_output(companies, extra_cols, output_file, input_df)

    elapsed      = time.time() - start
    mins, secs   = divmod(int(elapsed), 60)
    print(f"Færdig! Tid: {mins}m {secs}s")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_f  = sys.argv[1]
    output_f = sys.argv[2] if len(sys.argv) > 2 else "output.xlsx"
    main(input_f, output_f)

# CVR Lookup Tool

A desktop application that enriches a list of Danish companies with master data from the Danish Business Authority (Erhvervsstyrelsen) and structured financial KPIs from XBRL annual reports.

---

## Setup

### Prerequisites
No Python installation required. The tool ships as a self-contained Windows executable.

### Credentials
The tool connects to the Danish CVR Elasticsearch API, which requires credentials issued by Erhvervsstyrelsen.

Create a file named `.env` in the **same folder** as `cvr_lookup.exe`:

```
CVR_USER=your_username
CVR_PASSWORD=your_password
```

The tool will not start without valid credentials. If they are missing, the log will display the exact path where `.env` is expected.

---

## How to Use

1. Double-click `cvr_lookup.exe` to open the application.
2. Click **Vælg…** next to *Input fil* and select your Excel or CSV file.
3. Optionally change the output file path (defaults to `output.xlsx` in the same folder as the input).
4. Click **▶ Kør opslag** to start.
5. Progress is shown in the log window in real time.
6. When complete, click **📂 Åbn output** to open the result directly in Excel.

---

## Input File Format

The input file must be an `.xlsx`, `.xls`, or `.csv` file containing at least one of the following columns:

| Column | Description |
|---|---|
| `CVR` | 8-digit Danish CVR number — highest priority, fastest lookup |
| `Company` | Company name — searches CVR registry by name |
| `Contact person` | Person name — finds all companies where this person appears |
| `Contact person role` | Optional role filter used together with `Contact person` |

Additional columns in the input file (e.g. LinkedIn, Notes, Owner-manager likelihood) are passed through to the **Input** sheet in the output unchanged.

**Search priority:** If a row contains both a CVR number and a company name, the CVR number is used. If only a name is present, a name search is performed. If only a contact person is present, a person search is performed.

---

## How the Lookup Works

### Part 1 — Load input
The input file is read and validated. Each row becomes one lookup request.

### Part 2 — CVR master data
Each row is looked up against the CVR Elasticsearch API (`distribution.virk.dk`):

- **CVR search:** Exact match on CVR number. Returns exactly one company.
- **Company name search:** Full-text match on the current registered name. Returns up to 3,000 matches — for specific names this is typically 1–5.
- **Person search:** Searches for companies where the person's name appears as a participant (owner, director, board member), optionally filtered by role. Returns up to 3,000 matches.

If a name search returns multiple companies (e.g. "Cadesign Form" matches two separate legal entities), **each company gets its own row** in the output.

### Part 3 — Financial reports (XBRL)
For each company found, the tool fetches up to 4 years of annual reports from the public filings index. Reports must be submitted in **XBRL/XML format** to be parsed. The tool:

1. Queries the filings index for the most recent publications.
2. Downloads each XBRL file.
3. Identifies the main reporting period (longest duration context).
4. Extracts financial figures using a two-pass approach:
   - **Pass 1:** Builds a map of `(taxonomy, element_name) → value` for all elements in the main context.
   - **Pass 2:** Resolves named fields (revenue, EBIT, etc.) from `taxonomy_mappings.json`, trying the **FSA taxonomy first**, then **IFRS**.

### Part 4 — KPI calculations
Calculated from the fetched reports (see *Calculations* section below).

### Part 5 — Excel output
Results are written to a 6-sheet Excel workbook.

---

## Output Sheets

### Sheet 1 — Summary
One row per company. Key metrics and KPIs for quick screening.

### Sheet 2 — Company Info
All CVR master data: address, company type, status, industry, contact details, directors, board members, owners.

### Sheet 3 — Financial Data
One row per company per year. Structured income statement, balance sheet, and ratios.

### Sheet 4 — Financial Items
One row per company per year. All raw numeric XBRL elements found in the filing, labelled with their taxonomy prefix (`fsa:`, `ifrs:`). Useful for auditing which elements were present.

### Sheet 5 — Miscellaneous report items
Non-financial XBRL elements (text fields, metadata). Includes reporting class, audit statement type, etc.

### Sheet 6 — Input
The original input file, unchanged.

---

## Calculations

### EBITDA
```
EBITDA = EBIT + Depreciation and Amortisation (D&A)
```
EBIT and D&A are resolved from XBRL. If either is missing, EBITDA is blank.

### EBITDA Margin
```
EBITDA margin (%) = EBITDA / Revenue × 100
```

### EBIT Margin
```
EBIT margin (%) = EBIT / Revenue × 100
```

### Net Profit Margin
```
Net profit margin (%) = Net profit / Revenue × 100
```

### Equity Ratio
```
Equity ratio (%) = Equity / Total assets × 100
```

### Current Ratio
```
Current ratio = Current assets / Short-term liabilities
```
A ratio above 1 indicates the company can cover short-term obligations with current assets.

### Return on Equity (ROE)
```
ROE (%) = Net profit / Equity × 100
```

### Revenue CAGR / EBITDA CAGR / EBIT CAGR
Compound Annual Growth Rate between the oldest and newest available report years:
```
CAGR (%) = [(v_newest / v_oldest) ^ (1 / n) − 1] × 100
```
where `n` = number of years between the two data points.

**Returns N/A when:**
- Fewer than 2 years of data are available for the field.
- The oldest value is zero.
- The ratio `v_newest / v_oldest` is negative (sign change between years — mathematically produces a complex number and cannot be meaningfully interpreted as a growth rate).

The **CAGR period** column shows which years were used, e.g. `2021, 2024 (n=3)`.

### Rule of 40
```
Rule of 40 = Revenue CAGR (%) + EBITDA margin (%)
```
A combined score used to evaluate SaaS and growth companies. A score ≥ 40 is generally considered healthy. Returns N/A if either component is unavailable.

### Positive EBITDA Trend / Positive EBIT Trend
Determined by **linear regression** across all available years of EBITDA (or EBIT) data:

- A **positive slope** → `TRUE`
- A **negative slope** → `FALSE`
- Fewer than 2 data points → blank

Unlike CAGR, regression works regardless of whether values are negative, making it robust for companies in a turnaround phase. The **trend period** column lists all years included in the regression, e.g. `2021, 2022, 2023, 2024`.

### Positive Revenue Trend
```
TRUE if Revenue (latest year) > Revenue (oldest available year)
```
Simple endpoint comparison.

---

## Important Column Explanations

### EBIT elements used
The specific XBRL element name(s) from which EBIT was read. Example:
```
ProfitLossFromOrdinaryOperatingActivities
```
The FSA taxonomy is tried first. If the FSA element is not present, an IFRS element is used instead.

### No. of EBIT elements used
Always `1` — EBIT is always resolved from a single element (the first match in priority order).

### D&A elements used
The XBRL element name(s) from which Depreciation & Amortisation was read. May be a single element or a comma-separated list if two components were summed (e.g. property/plant/equipment depreciation + intangible asset amortisation).

### No. of D&A elements used
The number of XBRL elements that were **summed** to produce the D&A figure. `1` means a single combined element was found. `2` means two separate elements were added together.

### Report format
Whether the annual report was filed as `XBRL` (machine-readable, parseable by this tool) or `PDF` (human-readable only, no financial data extracted). Companies filing as PDF will appear in the output with blank financial columns.

### Reporting class
The Danish reporting class declared in the XBRL file:
| Class | Description |
|---|---|
| A | Micro — smallest companies, not required to file XBRL |
| B | Small — required to file XBRL |
| C | Medium/Large |
| D | Listed companies |

### Taxonomy version
The URL of the XBRL taxonomy schema used in the filing. Indicates whether the company used the FSA (Danish) taxonomy or IFRS, and which version year.

### Employees (latest)
An **interval string** from the CVR registry, e.g. `50-99`. This is the employment size band registered by the company — not a precise count.

### No. Employees (latest / -1y / -2y / -3y)
The **average number of employees** for the year as reported in the XBRL filing (`AverageNumberOfEmployees`). This is a precise figure, in contrast to the CVR interval above.

### Searched via / Matches found *(Company Info sheet)*
Records how the company was found (`CVR`, `Company`, or `Contact person`) and how many total matches the search returned. Useful for identifying ambiguous searches where multiple companies share a name.

---

## Known Limitations

- **PDF filings:** Companies that submit annual reports as PDF cannot be parsed. Financial columns will be blank; `Report format` will show `PDF`.
- **Reporting class A:** The smallest companies are not legally required to submit XBRL. Their filings may be absent or PDF only.
- **Name search ambiguity:** A company name search may return multiple legal entities. Each is included as a separate row — review `Matches found` and `CVR` to identify the intended company.
- **API size limit:** Each search returns a maximum of 3,000 results. For very generic names this may truncate results.
- **CAGR with sign changes:** If a company's EBITDA or EBIT crosses zero between the oldest and newest year, CAGR returns N/A. Use `Positive EBITDA Trend` (regression-based) instead.

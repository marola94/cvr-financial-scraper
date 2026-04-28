"""
Configuration — reads from .env file or environment variables.
Create a .env file based on .env.example before running.
"""
import os
import sys
from dotenv import load_dotenv

# Find .env next to the executable (PyInstaller) or next to this file (dev)
if getattr(sys, "frozen", False):
    _env_path = os.path.join(os.path.dirname(sys.executable), ".env")
else:
    _env_path = os.path.join(os.path.dirname(__file__), ".env")

load_dotenv(_env_path)

# ── API credentials ────────────────────────────────────────────────────────────
CVR_USER     = os.getenv("CVR_USER", "")
CVR_PASSWORD = os.getenv("CVR_PASSWORD", "")
CVR_AUTH     = (CVR_USER, CVR_PASSWORD)

# ── API endpoints ──────────────────────────────────────────────────────────────
CVR_SEARCH_URL      = "http://distribution.virk.dk/cvr-permanent/virksomhed/_search"
REGNSKAB_SEARCH_URL = "http://distribution.virk.dk/offentliggoerelser/_search"

# ── Behaviour ──────────────────────────────────────────────────────────────────
SLEEP_BETWEEN_CALLS = 0.3
YEARS_OF_FINANCIALS = 4   # Latest report + this many preceding years

# ── Input column names ─────────────────────────────────────────────────────────
# Priority order for CVR lookup when multiple are present
SEARCH_PRIORITY = ["CVR", "Company", "CPR", "Contact person"]
ROLE_COLUMN     = "Contact person role"

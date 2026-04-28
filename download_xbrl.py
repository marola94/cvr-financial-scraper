"""
Henter et XBRL-regnskab ned for et CVR-nummer og gemmer det som XML.
Kør: python download_xbrl.py
"""
import gzip
import xml.etree.ElementTree as ET
import requests
import config

CVR = "30919246"   # Skift til ønsket CVR-nummer
OUTPUT_FILE = "xbrl_regnskab.xml"

def find_xbrl_url(dokumenter):
    for dok in (dokumenter or []):
        if isinstance(dok, dict) and "xml" in dok.get("dokumentMimeType", "").lower():
            return dok.get("dokumentUrl")
    return None

query = {
    "query": {
        "bool": {
            "must": [
                {"term": {"cvrNummer": int(CVR)}},
                {"term": {"dokumenter.dokumentMimeType": "application"}},
                {"term": {"dokumenter.dokumentMimeType": "xml"}},
            ]
        }
    },
    "size": 5,
    "sort": [{"offentliggoerelsesTidspunkt": {"order": "desc"}}],
}

print(f"Søger regnskaber for CVR {CVR}...")
resp = requests.post(config.REGNSKAB_SEARCH_URL, auth=config.CVR_AUTH, json=query, timeout=30)
resp.raise_for_status()

hits = resp.json().get("hits", {}).get("hits", [])
print(f"Fandt {len(hits)} offentliggørelser")

for i, hit in enumerate(hits[:3]):
    src = hit.get("_source", {})
    xbrl_url = find_xbrl_url(src.get("dokumenter", []))
    dato = hit.get("_source", {}).get("offentliggoerelsesTidspunkt", "?")
    print(f"  [{i+1}] {dato} — {xbrl_url}")

# Hent det nyeste
xbrl_url = find_xbrl_url(hits[0].get("_source", {}).get("dokumenter", []))
if not xbrl_url:
    print("Ingen XBRL-URL fundet")
    exit(1)

print(f"\nHenter: {xbrl_url}")
xbrl_resp = requests.get(xbrl_url, timeout=30)
xbrl_resp.raise_for_status()

content = xbrl_resp.content
if not content.lstrip().startswith(b"<"):
    content = gzip.decompress(content)

# Smukkere XML-formatering
root = ET.fromstring(content)
ET.indent(root, space="  ")
tree = ET.ElementTree(root)
tree.write(OUTPUT_FILE, encoding="unicode", xml_declaration=True)

print(f"Gemt i: {OUTPUT_FILE}  ({len(content):,} bytes)")

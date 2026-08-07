import sys
import re
import json
import urllib.request
import urllib.error
import ssl
import openpyxl
from openpyxl.styles import Font, Alignment

# ==============================================================================
# 1. CONFIGURATION
# Keys (LeiKa IDs) used to query standard administrative services.
# ==============================================================================
LEIKAS_VERLUST = ['99008001014000', '99008001014001', '99008001014002']
LEIKAS_BEFREIUNG = ['99008002010000', '99008002010001', '99008002010003', '99008002010002']

# ==============================================================================
# 2. WEB CONNECTION HELPER
# Sends request to government portals over HTTPS directly using basic Python components.
# ==============================================================================
class DummyResponse:
    def __init__(self, status_code, text):
        self.status_code = status_code
        self.text = text

    def json(self):
        return json.loads(self.text)

def fetch_url(url, timeout=15):
    """Sends a request to a web URL and returns the text response."""
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

    # Skip SSL certificate verification to prevent local security software blocks
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as response:
            text = response.read().decode('utf-8')
            return DummyResponse(response.status, text)
    except urllib.error.HTTPError as e:
        try:
            text = e.read().decode('utf-8')
        except:
            text = ""
        return DummyResponse(e.code, text)
    except Exception as e:
        print(f"      [!] Netzwerkfehler bei {url}: {e}")
        return None

# ==============================================================================
# 3. SERVICE AVAILABILITY CHECK
# Checks whether the direct online portal links are functional for this ARS.
# ==============================================================================
def check_dienst_url(ars, service_type):
    if service_type == "Verlustmeldung":
        url = f"https://verwaltung.bund.de/onlinebeantragung/de/onlinedienst/1b723afb-c5f8-4ccd-82c5-d5f81afeda01/leistungsschluessel/99008001014002/region/{ars}"
        success_text = "Verlustmeldung Personalausweis"
    else:
        url = f"https://verwaltung.bund.de/onlinebeantragung/de/onlinedienst/bc032656-41c7-4357-980c-0d8e292feeb2/leistungsschluessel/99008002010000/region/{ars}"
        success_text = "Befreiung von der Ausweispflicht"

    resp = fetch_url(url)
    if resp:
        if success_text in resp.text:
            return "verfügbar"
        elif "Entschuldigung" in resp.text or "409" in str(resp.status_code):
            return "nicht verfügbar"
    return "nicht verfügbar"

# ==============================================================================
# 4. PVOG DATA RETRIEVAL
# Extracts Organization Unit IDs (OE) and Online Service IDs (OD) for a given ARS.
# ==============================================================================
def get_oe_od_ids(ars, leika):
    url = f"https://pvog.fitko.net/suchdienst/api/v3/servicedescriptions/leikaid?leikaIds={leika}&ars={ars}"
    resp = fetch_url(url)
    oe_ids, od_ids = [], []
    lbid = ""
    if resp and resp.status_code == 200:
        try:
            data = resp.json()
            if 'content' in data and data['content']:
                item = data['content'][0]
                lbid = item.get("id", "")
                s = json.dumps(item)
                # Search the text response for unique OE and OD identifiers
                oe_ids = list(dict.fromkeys(re.findall(r'([A-Za-z0-9]+\.OE\.[0-9]+)', s)))
                od_ids = list(dict.fromkeys(re.findall(r'([A-Za-z0-9]+\.OD\.[0-9]+)', s)))
        except Exception as e:
            print(f"      [!] Fehler beim Parsen von get_oe_od_ids: {e}")
    return oe_ids, od_ids, lbid

def check_oe_id(oe_id):
    """Retrieves the official name/title of an Organization Unit, excluding contact words."""
    if not oe_id: 
        return ""
    url = f"https://pvog.fitko.net/suchdienst/api/v5/organisationunits/detail?q={oe_id}"
    resp = fetch_url(url)
    title = ""
    if resp and resp.status_code == 200:
        try:
            data = resp.json()
            extracted_texts = []
            banned_words = [
                "telefon festnetz", "telefon", "e-mail", "fax", "telefax",
                "mobil", "mobiltelefon", "postanschrift", "besuchsadresse",
                "webseite", "internet", "hausanschrift"
            ]

            def extract_names(obj):
                if isinstance(obj, dict):
                    for k, v in obj.items():
                        if k in ["additionalInformation", "communicationSystems", "contactDetails", "contact", "contacts", "communications", "addresses", "channels", "communicationChannels", "openingHours", "paymentMethods"]:
                            continue
                        if k in ["title", "name"] and isinstance(v, str) and v.strip():
                            val = v.strip()
                            if val.lower() not in banned_words:
                                extracted_texts.append(val)
                        else:
                            extract_names(v)
                elif isinstance(obj, list):
                    for item in obj:
                        extract_names(item)

            extract_names(data)

            clean_texts = []
            for t in extracted_texts:
                if t not in clean_texts:
                    clean_texts.append(t)

            if len(clean_texts) >= 2:
                title = clean_texts[1]
            elif len(clean_texts) == 1:
                title = clean_texts[0]
        except Exception as e: 
            print(f"      [!] Fehler beim Parsen von check_oe_id ({oe_id}): {e}")
    return title

# ==============================================================================
# 5. ROUTING & SIGNATURE CHECKS
# Verifies route availability on FIT-Connect and digital signature status on XZuFi 2.3.
# ==============================================================================
def check_signatur(ars, leika):
    """Checks routing via the FIT-Connect system for all LeiKa IDs."""
    if not ars or not leika: 
        return "nicht vorhanden", ""
        
    url = f"https://routing-api-prod.fit-connect.fitko.net/v2/routes?leikaKey={leika}&ars={ars}"
        
    resp = fetch_url(url)
    if resp:
        try:
            raw_text = resp.text
            if "validation-exception" in raw_text:
                return "validation issue", ""
            if "constraint-violation" in raw_text:
                return "constraint violation", ""
                
            data = resp.json()
            count = data.get("count", 0)
            
            destination_name = ""
            routes = data.get("routes", [])
            if routes and len(routes) > 0:
                destination_name = routes[0].get("destinationName", "")

            if count == 1:
                return "vorhanden", destination_name
            elif count == 0:
                return "nicht vorhanden", destination_name
                
            if len(routes) > 0:
                return "vorhanden", destination_name
        except Exception as e:
            print(f"      [!] Fehler beim Parsen der Signatur für Leika {leika}: {e}")
            
    return "nicht vorhanden", ""

def check_signatur_oe(ars, leika, lbid=""):
    """Checks XZuFi 2.3 relation for digital signature and its length in PVOG for the two Leikas with the prod links."""
    # Restrict lookup strictly to primary targeted LeiKa IDs
    if leika not in ['99008001014002', '99008002010000']:
        return ""

    if not ars:
        return "nicht vorhanden"

    target_id = lbid if lbid else leika
    url = f"https://pvog.fitko.net/suchdienst/api/v1/relations/jzufi-2-3?ars={ars}&lbids={target_id}"

    resp = fetch_url(url)
    if resp and resp.status_code == 200:
        try:
            data = resp.json()
            values = []

            def find_id_sekundaer(obj):
                if isinstance(obj, dict):
                    for k, v in obj.items():
                        if k == "idSekundaer":
                            if isinstance(v, list):
                                for item in v:
                                    if isinstance(item, dict) and "value" in item:
                                        values.append(str(item["value"]).strip())
                            elif isinstance(v, dict) and "value" in v:
                                values.append(str(v["value"]).strip())
                        else:
                            find_id_sekundaer(v)
                elif isinstance(obj, list):
                    for item in obj:
                        find_id_sekundaer(item)

            find_id_sekundaer(data)

            if not values:
                return "nicht vorhanden"

            first_val = values[0]
            if len(first_val) < 500:
                return "evtl. zu kurz"
            else:
                return "vorhanden"

        except Exception as e:
            print(f"      [!] Fehler beim Parsen der OE-Signatur für Leika {leika}: {e}")

    return "nicht vorhanden"

def check_od_id(od_id):
    """Retrieves online service target URI and data protection link status."""
    if not od_id: return "", ""
    url = f"https://pvog.fitko.net/suchdienst/api/v2/onlineservices/detail?q={od_id}"
    resp = fetch_url(url)
    if resp and resp.status_code == 200:
        try:
            data = resp.json()
            links = data.get('links', [])
            uri = next((l.get('uri') for l in links if l.get('uri')), "")
            has_ds = "efa.datenschutzerklaerung.url" in json.dumps(data)
            return uri, "vorhanden" if has_ds else ""
        except Exception as e: 
            print(f"      [!] Fehler beim Parsen von check_od_id ({od_id}): {e}")
    return "", ""

# ==============================================================================
# 6. MAIN EXECUTION & EXCEL GENERATION
# Coordinates data gathering and formats the result into an Excel workbook.
# ==============================================================================
def main():
    # Prompt the user to enter the ARS number if not passed as a command-line argument
    if len(sys.argv) < 2:
        print("Willkommen beim PVOG Tool Annex Perso!")
        ars = input("Bitte geben Sie die 12-stellige ARS-Nummer ein und druecken Sie Enter: ").strip()
    else:
        ars = sys.argv[1].strip()

    if not ars:
        print("Fehler: Keine ARS-Nummer eingegeben. Programm wird beendet.")
        input("Druecken Sie Enter zum Schliessen...")
        sys.exit(1)

    output_file = f"PVOG_Resultate_{ars}.xlsx"

    print(f"\n[*] Starte Analyse fuer ARS: {ars}")

    # Check basic service access URLs
    print("[*] Pruefung der Produktivlinks...")
    status_v = check_dienst_url(ars, "Verlustmeldung")
    status_b = check_dienst_url(ars, "Befreiung")
    print(f"    - Verlustmeldung: {status_v}")
    print(f"    - Befreiung:      {status_b}")

    print(f"[*] Ausfuehrung PVOG Checks...")

    # Gather analysis rows into memory to decide whether secondary columns (OE2/OD2) are needed
    raw_results = []
    has_oe2 = False
    has_od2 = False

    services_to_check = [
        ("Verlustmeldung", LEIKAS_VERLUST),
        ("Befreiung", LEIKAS_BEFREIUNG)
    ]

    for dienst_name, leikas in services_to_check:
        for l_idx, leika in enumerate(leikas):
            oe_ids, od_ids, lbid = get_oe_od_ids(ars, leika)

            if len(oe_ids) > 1:
                has_oe2 = True
            if len(od_ids) > 1:
                has_od2 = True

            t1 = check_oe_id(oe_ids[0]) if len(oe_ids) > 0 else ""
            t2 = check_oe_id(oe_ids[1]) if len(oe_ids) > 1 else ""

            sig1, sig_name1 = check_signatur(ars, leika)
            sig_oe1 = check_signatur_oe(ars, leika, lbid)

            sig2, sig_name2 = check_signatur(ars, leika) if len(oe_ids) > 1 else ("", "")
            sig_oe2 = check_signatur_oe(ars, leika, lbid) if len(oe_ids) > 1 else "nicht vorhanden"

            uri1, ds1 = check_od_id(od_ids[0]) if len(od_ids) > 0 else ("", "")
            uri2, ds2 = check_od_id(od_ids[1]) if len(od_ids) > 1 else ("", "")

            row_info = {
                "dienst_name": dienst_name if l_idx == 0 else "",
                "leika": leika,
                "oe1": (oe_ids[0] if len(oe_ids) > 0 else "", t1, sig_name1, sig_oe1),
                "oe2": (oe_ids[1] if len(oe_ids) > 1 else "", t2, sig_name2, sig_oe2),
                "od1": (od_ids[0] if len(od_ids) > 0 else "", ds1, uri1),
                "od2": (od_ids[1] if len(od_ids) > 1 else "", ds2, uri2)
            }
            raw_results.append(row_info)

    # Build the header row dynamically without "OE-ID Signatur" & "OE-ID2 Signatur"
    headers = [
        "", "", "Dienst", "Leika",
        "OE-ID", "OE-ID Titel", "OE-ID Signatur Titel", "OE-ID Signatur bei Routing 2.3"
    ]

    if has_oe2:
        headers.extend([
            "OE-ID2", "OE-ID2 Titel", "OE-ID2 Signatur Titel", "OE-ID2 Signatur bei Routing 2.3"
        ])

    headers.extend([
        "OD-ID", "OD-ID Datenschutz-URL", "OD-ID URL"
    ])

    if has_od2:
        headers.extend([
            "OD-ID2", "OD-ID2 Datenschutz-URL", "OD-ID2 URL"
        ])

    # Build and write the Excel document
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PVOG Resultate"

        col_count = len(headers)
        ws.append([""] * col_count)
        ws.append(["", "1", "ARS", ars] + [""] * (col_count - 4))

        status_rows = [
            ["", "", f"Dienst-URL Verlustmeldung: {status_v}"] + [""] * (col_count - 3),
            ["", "", f"Dienst-URL Befreiung: {status_b}"] + [""] * (col_count - 3)
        ]
        for row in status_rows:
            ws.append(row)

        ws.append(headers)

        header_row_index = ws.max_row
        for col_idx in range(3, len(headers) + 1):
            cell = ws.cell(row=header_row_index, column=col_idx)
            cell.font = Font(bold=True)

        for r in raw_results:
            row_data = ["", "", r["dienst_name"], r["leika"]]
            row_data.extend(r["oe1"])
            
            if has_oe2:
                row_data.extend(r["oe2"])
                
            row_data.extend(r["od1"])
            
            if has_od2:
                row_data.extend(r["od2"])

            ws.append(row_data)

        # Automatically adjust Excel column widths so all contents fit clearly
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            if adjusted_width > 40:
                adjusted_width = 40
            ws.column_dimensions[column].width = adjusted_width

        wb.save(output_file)
        print(f"[*] Analyse erfolgreich abgeschlossen! Datei gespeichert: {output_file}")

    except Exception as main_err:
        print(f"\n[CRITICAL] Ein fataler Fehler verhinderte das Schreiben des Dokuments: {main_err}")

    if len(sys.argv) < 2:
        input("\nDruecken Sie Enter, um das Fenster zu schliessen...")

if __name__ == "__main__":
    main()
import streamlit as st
import json
import re
import ssl
import urllib.request
import urllib.error
import openpyxl

from io import BytesIO
from openpyxl.styles import Font, Alignment


# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="PVOG Annex Perso Prüftool",
    page_icon="🏛️",
    layout="wide"
)


# ============================================================
# LEIKA KONFIGURATION
# ============================================================

LEIKAS_VERLUST = [
    "99008001014000",
    "99008001014001",
    "99008001014002"
]


LEIKAS_BEFREIUNG = [
    "99008002010000",
    "99008002010001",
    "99008002010003",
    "99008002010002"
]


# ============================================================
# HTTP HELPER
# ============================================================

class DummyResponse:

    def __init__(self, status_code, text):
        self.status_code = status_code
        self.text = text


    def json(self):
        return json.loads(self.text)



def fetch_url(url, timeout=15):

    req = urllib.request.Request(
        url,
        headers={
            "User-Agent":
            "Mozilla/5.0 Streamlit PVOG Tool"
        }
    )


    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE


    try:

        with urllib.request.urlopen(
            req,
            timeout=timeout,
            context=ctx
        ) as response:

            text = response.read().decode(
                "utf-8"
            )

            return DummyResponse(
                response.status,
                text
            )


    except urllib.error.HTTPError as e:

        try:
            text = e.read().decode(
                "utf-8"
            )

        except:
            text = ""


        return DummyResponse(
            e.code,
            text
        )


    except Exception:

        return None



# ============================================================
# STREAMLIT STATUS LOGGER
# ============================================================

class StreamlitLogger:

    def __init__(self):

        self.placeholder = st.empty()
        self.messages = []


    def write(self, text):

        self.messages.append(text)

        self.placeholder.code(
            "\n".join(self.messages)
        )



# ============================================================
# DIENST URL CHECK
# ============================================================


def check_dienst_url(
        ars,
        service_type,
        logger=None
):

    if service_type == "Verlustmeldung":

        url = (
            "https://verwaltung.bund.de/"
            "onlinebeantragung/de/onlinedienst/"
            "1b723afb-c5f8-4ccd-82c5-d5f81afeda01/"
            f"leistungsschluessel/99008001014002/"
            f"region/{ars}"
        )

        success_text = (
            "Verlustmeldung Personalausweis"
        )


    else:

        url = (
            "https://verwaltung.bund.de/"
            "onlinebeantragung/de/onlinedienst/"
            "bc032656-41c7-4357-980c-0d8e292feeb2/"
            f"leistungsschluessel/99008002010000/"
            f"region/{ars}"
        )

        success_text = (
            "Befreiung von der Ausweispflicht"
        )


    if logger:
        logger.write(
            f"Prüfe Dienst URL: {service_type}"
        )


    resp = fetch_url(url)


    if resp:

        if success_text in resp.text:

            return "verfügbar"


        if (
            "Entschuldigung" in resp.text
            or "409" in str(resp.status_code)
        ):

            return "nicht verfügbar"


    return "nicht verfügbar"
# ============================================================
# PVOG OE / OD IDENTIFIKATION
# ============================================================


def get_oe_od_ids(
        ars,
        leika,
        logger=None
):

    url = (
        "https://pvog.fitko.net/"
        "suchdienst/api/v3/servicedescriptions/"
        f"leikaid?leikaIds={leika}&ars={ars}"
    )


    if logger:
        logger.write(
            f"PVOG Serviceabfrage: {leika}"
        )


    resp = fetch_url(url)


    oe_ids = []
    od_ids = []
    lbid = ""


    if resp and resp.status_code == 200:

        try:

            data = resp.json()


            if (
                "content" in data
                and data["content"]
            ):

                item = data["content"][0]

                lbid = item.get(
                    "id",
                    ""
                )


                text = json.dumps(
                    item
                )


                oe_ids = list(
                    dict.fromkeys(
                        re.findall(
                            r"([A-Za-z0-9]+\.OE\.[0-9]+)",
                            text
                        )
                    )
                )


                od_ids = list(
                    dict.fromkeys(
                        re.findall(
                            r"([A-Za-z0-9]+\.OD\.[0-9]+)",
                            text
                        )
                    )
                )


        except Exception as e:

            if logger:
                logger.write(
                    f"Fehler PVOG Parsing: {e}"
                )


    return (
        oe_ids,
        od_ids,
        lbid
    )



# ============================================================
# OE DETAILDATEN
# ============================================================


def check_oe_id(oe_id, logger=None):

    if not oe_id:
        return ""


    url = (
        "https://pvog.fitko.net/"
        "suchdienst/api/v5/"
        f"organisationunits/detail?q={oe_id}"
    )


    resp = fetch_url(url)

    title = ""


    if resp and resp.status_code == 200:

        try:

            data = resp.json()

            extracted_texts = []


            banned_words = [
                "telefon festnetz",
                "telefon",
                "e-mail",
                "fax",
                "telefax",
                "mobil",
                "mobiltelefon",
                "postanschrift",
                "besuchsadresse",
                "webseite",
                "internet",
                "hausanschrift"
            ]


            ignored_keys = [

                "additionalInformation",
                "communicationSystems",
                "contactDetails",
                "contact",
                "contacts",
                "communications",
                "addresses",
                "channels",
                "communicationChannels",
                "openingHours",
                "paymentMethods"

            ]


            def extract_names(obj):

                if isinstance(obj, dict):

                    for key, value in obj.items():


                        if key in ignored_keys:
                            continue


                        if key in [
                            "title",
                            "name"
                        ]:

                            if isinstance(value, str):

                                val = value.strip()


                                if (
                                    val
                                    and val.lower()
                                    not in banned_words
                                ):

                                    extracted_texts.append(
                                        val
                                    )

                        else:

                            extract_names(
                                value
                            )


                elif isinstance(obj, list):

                    for item in obj:

                        extract_names(
                            item
                        )


            extract_names(
                data
            )


            clean_texts = []


            for text in extracted_texts:

                if text not in clean_texts:

                    clean_texts.append(
                        text
                    )


            # gleiche Entscheidung wie Originalskript
            if len(clean_texts) >= 2:

                title = clean_texts[1]


            elif len(clean_texts) == 1:

                title = clean_texts[0]


        except Exception as e:

            if logger:

                logger.write(
                    f"OE Parsing Fehler {oe_id}: {e}"
                )


    return title


# ============================================================
# FIT CONNECT ROUTING
# ============================================================


def check_signatur(
        ars,
        leika,
        logger=None
):

    if not ars or not leika:

        return (
            "nicht vorhanden",
            ""
        )


    url = (
        "https://routing-api-prod."
        "fit-connect.fitko.net/v2/routes?"
        f"leikaKey={leika}&ars={ars}"
    )


    resp = fetch_url(url)


    if resp:

        try:

            raw = resp.text


            if "validation-exception" in raw:

                return (
                    "validation issue",
                    ""
                )


            data = resp.json()


            count = data.get(
                "count",
                0
            )


            routes = data.get(
                "routes",
                []
            )


            destination = ""


            if routes:

                destination = routes[0].get(
                    "destinationName",
                    ""
                )


            if count >= 1:

                return (
                    "vorhanden",
                    destination
                )


        except Exception:

            pass


    return (
        "nicht vorhanden",
        ""
    )



# ============================================================
# XZUFI 2.3 SIGNATUR
# ============================================================


def check_signatur_oe(
        ars,
        leika,
        lbid=""
):

    if leika not in [
        "99008001014002",
        "99008002010000"
    ]:

        return ""


    target = (
        lbid
        if lbid
        else leika
    )


    url = (
        "https://pvog.fitko.net/"
        "suchdienst/api/v1/"
        "relations/jzufi-2-3?"
        f"ars={ars}&lbids={target}"
    )


    resp = fetch_url(url)


    if resp and resp.status_code == 200:

        try:

            data = resp.json()

            values = []


            def find_values(obj):

                if isinstance(
                    obj,
                    dict
                ):

                    for key,value in obj.items():

                        if key == "idSekundaer":

                            if isinstance(
                                value,
                                list
                            ):

                                for x in value:

                                    if "value" in x:

                                        values.append(
                                            str(
                                                x["value"]
                                            )
                                        )

                        else:

                            find_values(
                                value
                            )


                elif isinstance(
                    obj,
                    list
                ):

                    for x in obj:

                        find_values(
                            x
                        )


            find_values(
                data
            )


            if not values:

                return (
                    "nicht vorhanden"
                )


            if len(values[0]) < 500:

                return (
                    "evtl. zu kurz"
                )


            return (
                "vorhanden"
            )


        except Exception:

            pass


    return (
        "nicht vorhanden"
    )

# ============================================================
# OD DETAILDATEN
# ============================================================


def check_od_id(
        od_id
):

    if not od_id:

        return (
            "",
            ""
        )


    url = (
        "https://pvog.fitko.net/"
        "suchdienst/api/v2/"
        f"onlineservices/detail?q={od_id}"
    )


    resp = fetch_url(url)


    uri = ""
    datenschutz = ""


    if resp and resp.status_code == 200:

        try:

            data = resp.json()


            links = data.get(
                "links",
                []
            )


            uri = next(
                (
                    x.get("uri")
                    for x in links
                    if x.get("uri")
                ),
                ""
            )


            if (
                "efa.datenschutzerklaerung.url"
                in json.dumps(data)
            ):

                datenschutz = "vorhanden"


        except Exception:

            pass


    return (
        uri,
        datenschutz
    )



# ============================================================
# HAUPTANALYSE
# ============================================================


def run_analysis(
        ars,
        logger
):


    logger.write(
        f"Starte Analyse für ARS {ars}"
    )


    status_verlust = check_dienst_url(
        ars,
        "Verlustmeldung",
        logger
    )


    status_befreiung = check_dienst_url(
        ars,
        "Befreiung",
        logger
    )


    logger.write(
        f"Verlustmeldung: {status_verlust}"
    )

    logger.write(
        f"Befreiung: {status_befreiung}"
    )


    raw_results = []

    has_oe2 = False
    has_od2 = False


    services = [

        (
            "Verlustmeldung",
            LEIKAS_VERLUST
        ),

        (
            "Befreiung",
            LEIKAS_BEFREIUNG
        )

    ]


    for service_name, leikas in services:


        for index, leika in enumerate(leikas):


            oe_ids, od_ids, lbid = get_oe_od_ids(
                ars,
                leika,
                logger
            )


            if len(oe_ids) > 1:

                has_oe2 = True


            if len(od_ids) > 1:

                has_od2 = True



            oe1 = (
                oe_ids[0]
                if len(oe_ids) > 0
                else ""
            )


            oe2 = (
                oe_ids[1]
                if len(oe_ids) > 1
                else ""
            )


            od1 = (
                od_ids[0]
                if len(od_ids) > 0
                else ""
            )


            od2 = (
                od_ids[1]
                if len(od_ids) > 1
                else ""
            )



            row = {


                "dienst":
                    service_name
                    if index == 0
                    else "",


                "leika":
                    leika,


                "oe1":
                [

                    oe1,

                    check_oe_id(
                        oe1
                    ),

                    check_signatur(
                        ars,
                        leika
                    )[1],

                    check_signatur_oe(
                        ars,
                        leika,
                        lbid
                    )

                ],



                "oe2":
                [

                    oe2,

                    check_oe_id(
                        oe2
                    ),

                    "",

                    ""

                ],



                "od1":
                [

                    od1,

                    check_od_id(
                        od1
                    )[1],

                    check_od_id(
                        od1
                    )[0]

                ],



                "od2":
                [

                    od2,

                    check_od_id(
                        od2
                    )[1],

                    check_od_id(
                        od2
                    )[0]

                ]

            }


            raw_results.append(
                row
            )


    return (
        raw_results,
        status_verlust,
        status_befreiung
    )



# ============================================================
# EXCEL ERZEUGUNG IM SPEICHER
# ============================================================


def create_excel(
        ars,
        results,
        status_v,
        status_b
):


    buffer = BytesIO()


    wb = openpyxl.Workbook()

    ws = wb.active

    ws.title = (
        "PVOG Resultate"
    )


    ws.append(
        [
            "",
            "ARS",
            ars
        ]
    )


    ws.append(
        [
            "",
            f"Dienst URL Verlustmeldung: {status_v}"
        ]
    )


    ws.append(
        [
            "",
            f"Dienst URL Befreiung: {status_b}"
        ]
    )



    headers = [

        "Dienst",
        "Leika",

        "OE-ID",
        "OE-ID Titel",
        "OE-ID Signatur Titel",
        "OE-ID Signatur Routing 2.3",

        "OE-ID2",
        "OE-ID2 Titel",

        "OD-ID",
        "OD-ID Datenschutz",
        "OD-ID URL",

        "OD-ID2",
        "OD-ID2 Datenschutz",
        "OD-ID2 URL"

    ]


    ws.append(
        headers
    )


    for cell in ws[4]:

        cell.font = Font(
            bold=True
        )



    for item in results:


        ws.append(

            [

                item["dienst"],

                item["leika"],


                *item["oe1"],


                *item["oe2"],


                *item["od1"],


                *item["od2"]

            ]

        )



    for column in ws.columns:

        max_length = 0

        letter = (
            column[0]
            .column_letter
        )


        for cell in column:

            if cell.value:

                max_length = max(
                    max_length,
                    len(
                        str(
                            cell.value
                        )
                    )
                )


        ws.column_dimensions[
            letter
        ].width = min(
            max_length + 2,
            45
        )


    wb.save(
        buffer
    )


    buffer.seek(
        0
    )


    return buffer



# ============================================================
# STREAMLIT BENUTZEROBERFLÄCHE
# ============================================================


st.title(
    "🏛️ PVOG Annex Perso Prüftool"
)


st.write(
"""
Dieses Tool prüft eine ARS-Region gegen:

- Verwaltungsportal Bund
- PVOG
- FIT-Connect Routing
- XZuFi 2.3
- Online-Dienste

und erzeugt eine Excel-Auswertung.
"""
)



ars = st.text_input(
    "12-stellige ARS Nummer",
    max_chars=12
)



if st.button(
    "🔍 Analyse starten"
):


    if len(ars) != 12:

        st.error(
            "Bitte eine gültige 12-stellige ARS eingeben."
        )


    else:


        logger = StreamlitLogger()


        with st.spinner(
            "Analyse läuft..."
        ):


            results, status_v, status_b = run_analysis(
                ars,
                logger
            )


            excel = create_excel(
                ars,
                results,
                status_v,
                status_b
            )


        st.success(
            "Analyse abgeschlossen."
        )


        st.download_button(

            label=
            "📥 Excel-Datei herunterladen",

            data=
            excel,

            file_name=
            f"PVOG_Resultate_{ars}.xlsx",

            mime=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )

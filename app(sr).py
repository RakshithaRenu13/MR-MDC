import os
import sqlite3
import hashlib
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st


# ============================================================
# EATON MDC SOLUTION CONFIGURATOR
# Compact Professional UI
# ============================================================

st.set_page_config(
    page_title="MDC Solution",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PATHS / CONSTANTS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MASTER_FILE = os.path.join(
    BASE_DIR,
    "MDC_Master_V1.xlsx"
)

TRACKING_DB = os.path.join(
    BASE_DIR,
    "MDC_Tracking.db"
)

DEMO_INTERNAL_PASSWORD = "MDC@123"


# ============================================================
# EATON COLOURS
# ============================================================

EATON_BLUE = "#007AC2"
EATON_DARK = "#004B91"

# Lighter ribbon colour for section headings
RIBBON_BLUE = "#3FA7D6"

EATON_LIGHT = "#EAF5FC"
BORDER = "#D6E4EE"

TEXT = "#263746"
MUTED = "#607486"


# ============================================================
# CONFIGURATION DISPLAY NAMES
# ============================================================

CONFIG_NAMES = {

    "Configuration 1":
        "1SR, 42U×800W×1200D, 3.5KW, W/O Dehumidifier",

    "Configuration 2":
        "1SR, 42U×800W×1200D, 7KW, Dehumidifier",

    "Configuration 3":
        "1SR, 42U×800W×1200D, 7KW, W/O Dehumidifier",

    "Configuration 4":
        "1SR, 42U×800W×1200D, 7KW, Dehumidifier",
}


# ============================================================
# GLOBAL CSS
# ============================================================

st.markdown(
    f"""
    <style>

    /* --------------------------------------------------------
       MAIN PAGE
       -------------------------------------------------------- */

    .block-container {{
        padding-top: 0.65rem;
        padding-bottom: 1.2rem;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 1500px;
    }}


    [data-testid="stHeader"] {{
        background: transparent;
    }}


    /* --------------------------------------------------------
       GENERAL FONT SIZE
       -------------------------------------------------------- */

    h1, h2, h3 {{
        color: {EATON_DARK};
        margin-top: 0.35rem !important;
        margin-bottom: 0.35rem !important;
    }}

    h2 {{
        font-size: 1.12rem !important;
    }}

    h3 {{
        font-size: 0.98rem !important;
    }}


    /* --------------------------------------------------------
       REDUCE STREAMLIT SPACING
       -------------------------------------------------------- */

    [data-testid="stVerticalBlock"] {{
        gap: 0.38rem;
    }}


    /* --------------------------------------------------------
       FORM LABELS
       -------------------------------------------------------- */

    div[data-testid="stTextInput"] label,
    div[data-testid="stTextArea"] label,
    div[data-testid="stSelectbox"] label,
    div[data-testid="stNumberInput"] label,
    div[data-testid="stRadio"] label {{
        font-size: 0.80rem !important;
        color: {MUTED} !important;
        font-weight: 600 !important;
    }}


    /* --------------------------------------------------------
       INPUT BOXES
       -------------------------------------------------------- */

    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input {{

        min-height: 34px !important;
        height: 34px !important;

        font-size: 0.86rem !important;
    }}


    div[data-baseweb="select"] > div {{

        min-height: 34px !important;

        font-size: 0.84rem !important;
    }}


    textarea {{

        font-size: 0.84rem !important;
    }}


    /* --------------------------------------------------------
       CHECKBOX
       -------------------------------------------------------- */

    div[data-testid="stCheckbox"] label p {{

        font-size: 0.82rem !important;
    }}


    /* --------------------------------------------------------
       RADIO
       -------------------------------------------------------- */

    div[data-testid="stRadio"] > div {{

        gap: 0.45rem !important;
    }}


    /* --------------------------------------------------------
       SECTION RIBBON
       -------------------------------------------------------- */

    .section-ribbon {{

        background: {RIBBON_BLUE};

        color: white;

        padding: 7px 12px;

        border-radius: 6px;

        font-size: 0.92rem;

        font-weight: 700;

        margin: 7px 0 5px 0;

        letter-spacing: 0.1px;
    }}


    /* --------------------------------------------------------
       MAIN TITLE
       -------------------------------------------------------- */

    .title-ribbon {{

        background: {EATON_BLUE};

        color: white;

        padding: 13px 18px;

        border-radius: 7px;

        margin-bottom: 8px;

        box-shadow:
            0 2px 7px rgba(0, 74, 145, 0.14);
    }}


    .title-main {{

        font-size: 1.55rem;

        font-weight: 750;

        line-height: 1.15;
    }}


    .title-sub {{

        font-size: 0.78rem;

        margin-top: 3px;

        opacity: 0.92;
    }}


    /* --------------------------------------------------------
       META CARDS
       -------------------------------------------------------- */

    .meta-card {{

        background: #F7FBFE;

        border: 1px solid {BORDER};

        border-radius: 6px;

        padding: 6px 10px;

        min-height: 48px;
    }}


    .meta-label {{

        color: {MUTED};

        font-size: 0.68rem;

        font-weight: 600;
    }}


    .meta-value {{

        color: {EATON_DARK};

        font-size: 0.82rem;

        font-weight: 700;

        margin-top: 2px;

        word-break: break-word;
    }}


    /* --------------------------------------------------------
       SELECTED CONFIGURATION
       -------------------------------------------------------- */

    .selection-card {{

        background: {EATON_LIGHT};

        border: 1px solid #B9DDF2;

        border-left: 4px solid {EATON_BLUE};

        border-radius: 6px;

        padding: 7px 11px;

        font-size: 0.80rem;

        color: {EATON_DARK};

        font-weight: 650;

        margin: 2px 0 4px 0;
    }}


    /* --------------------------------------------------------
       PRICE CARDS
       -------------------------------------------------------- */

    .price-card {{

        background: #F8FAFC;

        border: 1px solid {BORDER};

        border-radius: 6px;

        padding: 8px 10px;

        min-height: 61px;
    }}


    .price-label {{

        color: {MUTED};

        font-size: 0.69rem;

        font-weight: 600;
    }}


    .price-value {{

        color: {EATON_DARK};

        font-size: 1.03rem;

        font-weight: 750;

        margin-top: 2px;

        white-space: nowrap;
    }}


    /* --------------------------------------------------------
       FINAL PRICE
       -------------------------------------------------------- */

    .final-price-card {{

        background: {EATON_LIGHT};

        border: 1px solid #9ED0EF;

        border-left: 5px solid {EATON_BLUE};

        border-radius: 7px;

        padding: 9px 13px;

        min-height: 61px;
    }}


    .final-price-label {{

        color: {EATON_DARK};

        font-size: 0.72rem;

        font-weight: 700;
    }}


    .final-price-value {{

        color: {EATON_DARK};

        font-size: 1.18rem;

        font-weight: 800;

        margin-top: 2px;

        white-space: nowrap;
    }}


    /* --------------------------------------------------------
       SMALL NOTES
       -------------------------------------------------------- */

    .compact-note {{

        color: {MUTED};

        font-size: 0.70rem;

        margin: 1px 0 3px 0;
    }}


    /* --------------------------------------------------------
       FINAL BOQ TABLE
       -------------------------------------------------------- */

    .final-boq-table {{

        width: 100%;

        border-collapse: collapse;

        border: 1px solid {BORDER};

        border-radius: 6px;

        overflow: hidden;

        font-family: Arial, sans-serif;

        font-size: 0.76rem;
    }}


    .final-boq-table th {{

        background: #F1F6FA;

        color: {EATON_DARK};

        font-weight: 700;

        padding: 7px 7px;

        border-bottom: 1px solid {BORDER};

        text-align: left;
    }}


    .final-boq-table td {{

        padding: 6px 7px;

        border-bottom: 1px solid #E7EEF3;

        color: {TEXT};

        vertical-align: middle;
    }}


    .final-boq-table .center {{

        text-align: center;
    }}


    /* --------------------------------------------------------
       MAIN MDC TITLE ROW
       -------------------------------------------------------- */

    .main-mdc-row td {{

        background: {EATON_DARK};

        color: white !important;

        font-weight: 700;

        text-align: center !important;

        padding: 8px;
    }}


    /* --------------------------------------------------------
       BOQ SECTION ROW
       -------------------------------------------------------- */

    .section-row td {{

        background: {EATON_BLUE};

        color: white !important;

        font-weight: 700;

        text-align: center !important;

        padding: 6px;
    }}


    /* --------------------------------------------------------
       HISTORY
       -------------------------------------------------------- */

    .history-table {{

        width: 100%;

        border-collapse: collapse;

        font-size: 0.76rem;
    }}


    .history-table th {{

        background: #F1F6FA;

        color: {EATON_DARK};

        padding: 7px;

        border: 1px solid {BORDER};

        text-align: left;
    }}


    .history-table td {{

        padding: 6px 7px;

        border: 1px solid {BORDER};
    }}


    /* --------------------------------------------------------
       BUTTONS
       -------------------------------------------------------- */

    div.stButton > button,
    div.stDownloadButton > button {{

        min-height: 34px;

        height: 34px;

        font-size: 0.80rem;

        font-weight: 650;

        border-radius: 5px;
    }}


    .stCaption {{

        font-size: 0.70rem !important;
    }}


    hr {{

        margin: 0.5rem 0 !important;
    }}

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DATABASE
# ============================================================

def init_tracking_db():

    conn = sqlite3.connect(TRACKING_DB)

    cur = conn.cursor()

    # --------------------------------------------------------
    # Configuration table
    # --------------------------------------------------------

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS configurations (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            configuration_id TEXT UNIQUE,

            created_at TEXT,

            customer_name TEXT,

            customer_place TEXT,

            problem TEXT,

            solution TEXT,

            mdc_type TEXT,

            configuration TEXT,

            base_cost REAL,

            optional_cost REAL,

            pdu_cost REAL,

            total_cost REAL,

            margin_pct REAL,

            freight REAL,

            installation REAL,

            warranty_pct REAL,

            margin_price REAL,

            final_selling_price REAL,

            warranty_amount REAL
        )
        """
    )

    # --------------------------------------------------------
    # Configuration items
    # --------------------------------------------------------

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS configuration_items (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            configuration_id TEXT,

            component_type TEXT,

            part_code TEXT,

            description TEXT,

            quantity REAL,

            uom TEXT,

            unit_cost REAL,

            total_cost REAL,

            unit_price REAL,

            total_price REAL
        )
        """
    )

    # --------------------------------------------------------
    # Download history
    #
    # User Count is based on Excel/PDF downloads.
    # --------------------------------------------------------

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS download_history (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            user_code TEXT,

            downloaded_at TEXT,

            customer_name TEXT,

            mdc_type TEXT,

            configuration TEXT,

            file_type TEXT
        )
        """
    )

    conn.commit()

    conn.close()


def get_download_count():

    conn = sqlite3.connect(TRACKING_DB)

    cur = conn.cursor()

    cur.execute(
        "SELECT COUNT(*) FROM download_history"
    )

    count = cur.fetchone()[0]

    conn.close()

    return int(count)


def record_download(
    user_code,
    file_type
):

    conn = sqlite3.connect(TRACKING_DB)

    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO download_history
        (
            user_code,
            downloaded_at,
            customer_name,
            mdc_type,
            configuration,
            file_type
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            user_code,

            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

            st.session_state.customer_name,

            st.session_state.mdc_type,

            st.session_state.configuration,

            file_type,
        ),
    )

    conn.commit()

    conn.close()


def generate_configuration_id():

    today = datetime.now().strftime(
        "%Y%m%d"
    )

    conn = sqlite3.connect(TRACKING_DB)

    cur = conn.cursor()

    cur.execute(
        """
        SELECT COUNT(*)
        FROM configurations
        WHERE configuration_id LIKE ?
        """,
        (
            f"MDC-{today}-%",
        ),
    )

    count = cur.fetchone()[0] + 1

    conn.close()

    return (
        f"MDC-{today}-{count:04d}"
    )


init_tracking_db()


# ============================================================
# MASTER DATA
# ============================================================

@st.cache_data
def load_master():

    configs = pd.read_excel(
        MASTER_FILE,
        sheet_name="Configurations"
    )

    components = pd.read_excel(
        MASTER_FILE,
        sheet_name="Components"
    )

    accessories = pd.read_excel(
        MASTER_FILE,
        sheet_name="Accessories"
    )

    pdus = pd.read_excel(
        MASTER_FILE,
        sheet_name="PDUs"
    )

    return (
        configs,
        components,
        accessories,
        pdus
    )


configs_df, components_df, accessories_df, pdus_df = load_master()


# ============================================================
# SESSION STATE
# ============================================================

defaults = {

    "mode": "Sales",

    "authenticated": False,

    "customer_name": "",

    "customer_place": "",

    "problem": "",

    "solution": "",

    "mdc_type": "Single Rack",

    "configuration": "Configuration 1",

    "accessory_qty": {},

    "pdu_qty": {},

    "pdu_type": "Basic",

    "pdu_selected_part": "None",

    "fire_type": "None",

    "camera_enabled": False,

    "configuration_id": None,

    "margin_pct": 20.0,

    "freight": 0.0,

    "installation": 0.0,

    "warranty_pct": 0.0,
}


for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value


if st.session_state.configuration_id is None:

    st.session_state.configuration_id = (
        generate_configuration_id()
    )


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def money(value):

    return f"₹ {float(value):,.2f}"


def price_box(
    label,
    value,
    final=False
):

    cls = (
        "final-price-card"
        if final
        else "price-card"
    )

    label_cls = (
        "final-price-label"
        if final
        else "price-label"
    )

    value_cls = (
        "final-price-value"
        if final
        else "price-value"
    )

    st.markdown(
        f"""
        <div class="{cls}">

            <div class="{label_cls}">
                {label}
            </div>

            <div class="{value_cls}">
                {money(value)}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


def ribbon(text):

    st.markdown(
        f"""
        <div class="section-ribbon">
            {text}
        </div>
        """,
        unsafe_allow_html=True,
    )


def internal_password():

    try:

        return st.secrets[
            "MDC_INTERNAL_PASSWORD"
        ]

    except Exception:

        return DEMO_INTERNAL_PASSWORD


def selected_config_record():

    match = configs_df[
        (
            configs_df["MDC Type"]
            == st.session_state.mdc_type
        )
        &
        (
            configs_df["Configuration"]
            == st.session_state.configuration
        )
    ]

    if not match.empty:

        return match.iloc[0]

    return None


def selected_components():

    return components_df[
        (
            components_df["MDC Type"]
            == st.session_state.mdc_type
        )
        &
        (
            components_df["Configuration"]
            == st.session_state.configuration
        )
    ].copy()


def clean_part(value):

    value = str(value).strip()

    if value.lower() == "nan":

        return ""

    return value


# ============================================================
# USER CODE
#
# Format:
#
# SR-C1-AB12-0001
#
# SR = Single Rack
# MR = Multirack
# C1 = Configuration
# AB12 = component/configuration selection fingerprint
# 0001 = download/user count
#
# Thus the code changes according to the selected
# configuration/components and each download gets a
# separate sequence number.
# ============================================================

def option_signature():

    parts = [

        st.session_state.mdc_type,

        st.session_state.configuration,

        st.session_state.get(
            "pdu_type",
            "None"
        ),

        st.session_state.get(
            "pdu_selected_part",
            "None"
        ),

        st.session_state.get(
            "fire_type",
            "None"
        ),

        (
            "CAM1"
            if st.session_state.get(
                "camera_enabled",
                False
            )
            else "CAM0"
        ),
    ]

    for part, qty in sorted(
        st.session_state.accessory_qty.items()
    ):

        if float(qty) > 0:

            parts.append(
                f"{part}:{qty}"
            )

    raw = "|".join(parts).encode(
        "utf-8"
    )

    return hashlib.sha1(
        raw
    ).hexdigest()[:4].upper()


def make_user_code(sequence):

    type_code = (
        "SR"
        if st.session_state.mdc_type
        == "Single Rack"
        else "MR"
    )

    cfg_number = "C1"

    text = str(
        st.session_state.configuration
    )

    digits = "".join(
        ch
        for ch in text
        if ch.isdigit()
    )

    if digits:

        cfg_number = f"C{digits}"

    return (
        f"{type_code}-"
        f"{cfg_number}-"
        f"{option_signature()}-"
        f"{sequence:04d}"
    )


def current_user_code():

    return make_user_code(
        get_download_count() + 1
    )


# ============================================================
# BUILD BOM
# ============================================================

def build_bom():

    rows = []

    # --------------------------------------------------------
    # Base configuration
    # --------------------------------------------------------

    for _, r in selected_components().iterrows():

        cost = r["Unit Cost"]

        qty = float(
            r["Quantity"]
        )

        rows.append(
            {

                "S.No.": len(rows) + 1,

                "Component Type":
                    "Base (Configuration)",

                "Part Code":
                    clean_part(
                        r["Part Code"]
                    ),

                "Description":
                    r["Description"],

                "Quantity":
                    qty,

                "UOM":
                    r["UOM"],

                "Unit Cost":
                    cost,

                "Total Cost":
                    (
                        cost * qty
                        if pd.notna(cost)
                        else None
                    ),

                "Source":
                    "Configuration",
            }
        )


    # --------------------------------------------------------
    # Optional accessories
    # --------------------------------------------------------

    for _, r in accessories_df.iterrows():

        part = clean_part(
            r["Part Code"]
        )

        qty = float(
            st.session_state
            .accessory_qty
            .get(part, 0)
        )

        if qty > 0:

            cost = r["Unit Cost"]

            rows.append(
                {

                    "S.No.":
                        len(rows) + 1,

                    "Component Type":
                        "Optional Accessory",

                    "Part Code":
                        part,

                    "Description":
                        r["Description"],

                    "Quantity":
                        qty,

                    "UOM":
                        r["UOM"],

                    "Unit Cost":
                        cost,

                    "Total Cost":
                        (
                            cost * qty
                            if pd.notna(cost)
                            else None
                        ),

                    "Source":
                        "Optional Accessory",
                }
            )


    # --------------------------------------------------------
    # PDU
    # --------------------------------------------------------

    for _, r in pdus_df.iterrows():

        part = clean_part(
            r["Part Code"]
        )

        qty = float(
            st.session_state
            .pdu_qty
            .get(part, 0)
        )

        if qty > 0:

            cost = r["Unit Cost"]

            desc = (
                f'{r["Description"]} | '
                f'Type: {r["Type"]} | '
                f'C13: {r["C13"]} | '
                f'C19: {r["C19"]}'
            )

            rows.append(
                {

                    "S.No.":
                        len(rows) + 1,

                    "Component Type":
                        "PDU",

                    "Part Code":
                        part,

                    "Description":
                        desc,

                    "Quantity":
                        qty,

                    "UOM":
                        r["UOM"],

                    "Unit Cost":
                        cost,

                    "Total Cost":
                        (
                            cost * qty
                            if pd.notna(cost)
                            else None
                        ),

                    "Source":
                        "PDU",
                }
            )


    return pd.DataFrame(rows)


# ============================================================
# COST SUMMARY
# ============================================================

def cost_summary(bom):

    cfg = selected_config_record()

    if (
        cfg is not None
        and pd.notna(cfg["Base Cost"])
    ):

        base_cost = float(
            cfg["Base Cost"]
        )

    else:

        base_cost = 0.0


    optional_cost = 0.0

    pdu_cost = 0.0


    if not bom.empty:

        optional_cost = float(
            bom.loc[
                bom["Source"]
                == "Optional Accessory",
                "Total Cost"
            ]
            .fillna(0)
            .sum()
        )


        pdu_cost = float(
            bom.loc[
                bom["Source"]
                == "PDU",
                "Total Cost"
            ]
            .fillna(0)
            .sum()
        )


    total_cost = (
        base_cost
        + optional_cost
        + pdu_cost
    )

    return (
        base_cost,
        optional_cost,
        pdu_cost,
        total_cost
    )


# ============================================================
# SELLING PRICE
# ============================================================

def add_selling_prices(
    bom,
    total_cost,
    margin_pct,
    freight,
    installation
):

    result = bom.copy()

    margin_price = (
        total_cost
        / (1 - margin_pct / 100)
        if margin_pct < 100
        else 0
    )

    final_selling_price = (
        margin_price
        + freight
        + installation
    )


    known_cost_total = (

        result["Total Cost"]
        .fillna(0)
        .sum()

        if not result.empty

        else 0
    )


    if known_cost_total > 0:

        result["Total Price"] = (

            result["Total Cost"]
            .fillna(0)
            / known_cost_total
            * final_selling_price
        )

        result["Unit Price"] = (

            result["Total Price"]
            / result["Quantity"]
        )

    else:

        result["Total Price"] = pd.NA

        result["Unit Price"] = pd.NA


    return (
        result,
        margin_price,
        final_selling_price
    )


# ============================================================
# USER ACCESS
# ============================================================

with st.sidebar:

    st.markdown(
        "### User Access"
    )

    mode = st.radio(
        "Select User Type",
        [
            "Sales",
            "Internal – MDC"
        ],
        index=(
            0
            if st.session_state.mode
            == "Sales"
            else 1
        ),
    )


    if mode != st.session_state.mode:

        st.session_state.mode = mode

        if mode == "Sales":

            st.session_state.authenticated = False

        st.rerun()


    if mode == "Internal – MDC":

        if not st.session_state.authenticated:

            pwd = st.text_input(
                "MDC Password",
                type="password"
            )

            if st.button(
                "Unlock Internal Mode",
                use_container_width=True
            ):

                if (
                    pwd
                    == internal_password()
                ):

                    st.session_state.authenticated = True

                    st.rerun()

                else:

                    st.error(
                        "Incorrect password."
                    )

        else:

            st.success(
                "Internal mode unlocked."
            )

            if st.button(
                "Lock Internal Mode",
                use_container_width=True
            ):

                st.session_state.authenticated = False

                st.session_state.mode = "Sales"

                st.rerun()


is_internal = (

    st.session_state.mode
    == "Internal – MDC"

    and
    st.session_state.authenticated
)


# ============================================================
# TITLE
# ============================================================

st.markdown(
    """
    <div class="title-ribbon">

        <div class="title-main">
            Eaton MDC Solution Configurator
        </div>

        <div class="title-sub">
            Modular Data Center Solution Configuration &amp; Pricing
        </div>

    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# USER CODE / USER COUNT / DATE
# ============================================================

download_count = get_download_count()

next_code = current_user_code()

current_date = datetime.now().strftime(
    "%d-%m-%Y"
)


m1, m2, m3, m4 = st.columns(
    [
        1.45,
        1.0,
        1.0,
        1.0
    ]
)


with m1:

    st.markdown(
        f"""
        <div class="meta-card">

            <div class="meta-label">
                User Code
            </div>

            <div class="meta-value">
                {next_code}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


with m2:

    st.markdown(
        f"""
        <div class="meta-card">

            <div class="meta-label">
                User Count
            </div>

            <div class="meta-value">
                {download_count}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


with m3:

    st.markdown(
        f"""
        <div class="meta-card">

            <div class="meta-label">
                Date
            </div>

            <div class="meta-value">
                {current_date}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


with m4:

    st.markdown(
        f"""
        <div class="meta-card">

            <div class="meta-label">
                Access
            </div>

            <div class="meta-value">
                {"Internal" if is_internal else "Sales"}
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# CUSTOMER DETAILS
# NOT NUMBERED - COMPACT TOP AREA
# ============================================================

st.markdown(
    "**Customer Details**",
    unsafe_allow_html=True
)


c1, c2 = st.columns(
    [1.15, 1.0]
)


with c1:

    st.text_input(
        "Customer Name *",
        key="customer_name",
        placeholder="Enter customer name",
    )


with c2:

    st.text_input(
        "Customer Place",
        key="customer_place",
        placeholder="Enter location",
    )


p1, p2 = st.columns(2)


with p1:

    st.text_area(
        "Problem Description",
        key="problem",
        height=54,
        placeholder="Optional",
    )


with p2:

    st.text_area(
        "Solution",
        key="solution",
        height=54,
        placeholder="Optional",
    )


# ============================================================
# 1. MDC TYPE & CONFIGURATION
# ============================================================

ribbon(
    "1. MDC Type & Configuration"
)


mtype_col, config_col = st.columns(
    [1.0, 2.0]
)


with mtype_col:

    mdc_type = st.selectbox(
        "MDC Type",
        [
            "Single Rack",
            "Multirack"
        ],
        index=(
            0
            if st.session_state.mdc_type
            == "Single Rack"
            else 1
        ),
        key="mdc_type_select",
    )


if (
    mdc_type
    != st.session_state.mdc_type
):

    st.session_state.mdc_type = mdc_type

    st.session_state.configuration = (
        "Configuration 1"
    )

    st.session_state.accessory_qty = {}

    st.session_state.pdu_qty = {}

    st.session_state.fire_type = "None"

    st.session_state.camera_enabled = False

    st.session_state.configuration_id = (
        generate_configuration_id()
    )

    st.rerun()


available = configs_df[
    configs_df["MDC Type"]
    == st.session_state.mdc_type
].copy()


labels = available[
    "Configuration"
].tolist()


with config_col:

    if labels:

        selected_config = st.selectbox(
            "Configuration",
            labels,

            index=(
                labels.index(
                    st.session_state.configuration
                )
                if
                st.session_state.configuration
                in labels
                else 0
            ),

            format_func=lambda x:
                CONFIG_NAMES.get(
                    x,
                    x
                ),

            key="configuration_select",
        )


        if (
            selected_config
            != st.session_state.configuration
        ):

            st.session_state.configuration = (
                selected_config
            )

            st.session_state.configuration_id = (
                generate_configuration_id()
            )

            st.session_state.pdu_qty = {}

            st.session_state.pdu_selected_part = (
                "None"
            )

            st.rerun()

    else:

        st.session_state.configuration = (
            "Configuration 1"
        )

        st.info(
            "No configuration data available."
        )


# ------------------------------------------------------------
# Selected Configuration
# ------------------------------------------------------------

cfg = selected_config_record()


if cfg is not None:

    display_name = CONFIG_NAMES.get(

        st.session_state.configuration,

        str(
            cfg.get(
                "Configuration Title",
                st.session_state.configuration
            )
        )
    )


    st.markdown(
        f"""
        <div class="selection-card">

            Selected Configuration:
            {display_name}

        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# 2. PDU SELECTION
# ============================================================

ribbon(
    "2. PDU Selection"
)


pdu_type = st.radio(
    "PDU Type",
    [
        "Basic",
        "Metered",
        "Switched",
        "None"
    ],

    horizontal=True,

    index=[
        "Basic",
        "Metered",
        "Switched",
        "None"
    ].index(
        st.session_state.get(
            "pdu_type",
            "Basic"
        )
    ),

    key="pdu_type_radio",
)


if (
    pdu_type
    != st.session_state.pdu_type
):

    st.session_state.pdu_type = pdu_type

    st.session_state.pdu_qty = {}

    st.session_state.pdu_selected_part = (
        "None"
    )

    st.rerun()


# ------------------------------------------------------------
# Filter PDU components from Excel
# ------------------------------------------------------------

pdu_filtered = pdus_df.copy()


if (
    pdu_type != "None"
    and "Type" in pdu_filtered.columns
):

    pdu_filtered = pdu_filtered[
        pdu_filtered["Type"]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        pdu_type.lower()
    ].copy()


if pdu_type == "None":

    st.session_state.pdu_qty = {}

    st.session_state.pdu_selected_part = (
        "None"
    )

else:

    pdu_labels = [

        f'{clean_part(r["Part Code"])} — '
        f'{r["Description"]}'

        for _, r
        in pdu_filtered.iterrows()
    ]


    pc1, pc2 = st.columns(
        [5.5, 1.1]
    )


    with pc1:

        if pdu_labels:

            selected_pdu = st.selectbox(

                "Select PDU",

                ["None"] + pdu_labels,

                index=(

                    (
                        ["None"] + pdu_labels
                    ).index(
                        st.session_state
                        .pdu_selected_part
                    )

                    if
                    st.session_state
                    .pdu_selected_part
                    in
                    (
                        ["None"] + pdu_labels
                    )

                    else 0
                ),

                key="pdu_dropdown",
            )

        else:

            selected_pdu = "None"

            st.selectbox(
                "Select PDU",
                [
                    "No PDU components available"
                ],
                disabled=True
            )


    with pc2:

        qty = st.number_input(
            "Qty",

            min_value=1,

            max_value=999,

            value=1,

            step=1,

            key="pdu_quantity",
        )


    st.session_state.pdu_qty = {}

    st.session_state.pdu_selected_part = (
        selected_pdu
    )


    if selected_pdu != "None":

        selected_index = (
            pdu_labels.index(
                selected_pdu
            )
        )

        selected_row = (
            pdu_filtered.iloc[
                selected_index
            ]
        )

        part = clean_part(
            selected_row["Part Code"]
        )

        st.session_state.pdu_qty[
            part
        ] = qty


# ============================================================
# 3. OTHER ACCESSORIES
# ============================================================

ribbon(
    "3. Other Accessories"
)


optional_lookup = {

    clean_part(
        r["Part Code"]
    ): r

    for _, r
    in accessories_df.iterrows()

    if clean_part(
        r["Part Code"]
    )
}


# ------------------------------------------------------------
# PART CODES
# ------------------------------------------------------------

FIRE_SUPPRESSION_PARTS = [

    "801073203",

    "HRD-XH1C",
]


FIRE_SUPPRESSION_SUPPORT = [

    "801075235",
]


CAMERA_PARTS = [

    "801303201",

    "801303202",

    "801303204",

    "801303206",

    "801303208",

    "801303203",
]


OTHER_OPTIONAL_PARTS = [

    "801223664",

    "801075237",

    "801029022",
]


# ============================================================
# 3.1 FIRE SUPPRESSION
# ============================================================

fa, fb = st.columns(
    [1.0, 2.2]
)


with fa:

    st.markdown(
        "**3.1 Fire Suppression**"
    )


with fb:

    fire_options = [
        "None",
        "External",
        "Internal"
    ]

    fire_type = st.radio(

        "Fire Suppression Type",

        fire_options,

        horizontal=True,

        label_visibility="collapsed",

        index=fire_options.index(
            st.session_state.fire_type
        ),

        key="fire_type_radio",
    )

    st.session_state.fire_type = (
        fire_type
    )


# ------------------------------------------------------------
# Clear existing fire selections
# ------------------------------------------------------------

for part in (

    FIRE_SUPPRESSION_PARTS
    +
    FIRE_SUPPRESSION_SUPPORT

):

    st.session_state.accessory_qty.pop(
        part,
        None
    )


# ------------------------------------------------------------
# External
# ------------------------------------------------------------

if fire_type == "External":

    part = FIRE_SUPPRESSION_PARTS[0]

    if part in optional_lookup:

        st.session_state.accessory_qty[
            part
        ] = 1


    for support in FIRE_SUPPRESSION_SUPPORT:

        if support in optional_lookup:

            st.session_state.accessory_qty[
                support
            ] = 1


# ------------------------------------------------------------
# Internal
# ------------------------------------------------------------

elif fire_type == "Internal":

    part = FIRE_SUPPRESSION_PARTS[1]

    if part in optional_lookup:

        st.session_state.accessory_qty[
            part
        ] = 1


    for support in FIRE_SUPPRESSION_SUPPORT:

        if support in optional_lookup:

            st.session_state.accessory_qty[
                support
            ] = 1


# ============================================================
# 3.2 CAMERA
# ============================================================

ca, cb = st.columns(
    [1.0, 2.2]
)


with ca:

    st.markdown(
        "**3.2 Camera**"
    )


with cb:

    camera = st.radio(

        "Camera System",

        [
            "No",
            "Yes"
        ],

        horizontal=True,

        index=(
            1
            if st.session_state
            .camera_enabled
            else 0
        ),

        label_visibility="collapsed",

        key="camera_radio",
    )


    st.session_state.camera_enabled = (
        camera == "Yes"
    )


# ------------------------------------------------------------
# Camera components
# ------------------------------------------------------------

for part in CAMERA_PARTS:

    st.session_state.accessory_qty.pop(
        part,
        None
    )


if st.session_state.camera_enabled:

    for part in CAMERA_PARTS:

        if part in optional_lookup:

            st.session_state.accessory_qty[
                part
            ] = 1


# ============================================================
# 3.3 - 3.6 OTHER ACCESSORIES
# ============================================================

accessory_labels = [

    (
        "3.3",
        "Rotating Keyboard",
        "801223664"
    ),

    (
        "3.4",
        "Cable Manager",
        "801075237"
    ),

    (
        "3.5",
        "Top Cable Tray",
        "801029022"
    ),

    (
        "3.6",
        "Brush Panel",
        "801075235"
    ),
]


for number, label, part in accessory_labels:

    if part not in optional_lookup:

        continue


    # --------------------------------------------------------
    # Brush Panel is automatically controlled
    # by Fire Suppression.
    # --------------------------------------------------------

    if part == "801075235":

        selected = (

            st.session_state
            .accessory_qty
            .get(part, 0)
            > 0
        )


        x1, x2 = st.columns(
            [4.6, 1.0]
        )


        with x1:

            st.checkbox(

                f"{number} {label}",

                value=selected,

                disabled=True,

                key=f"fixed_{part}",
            )


        with x2:

            st.caption("Auto")

        continue


    # --------------------------------------------------------
    # Normal accessory
    # --------------------------------------------------------

    x1, x2 = st.columns(
        [4.6, 1.0]
    )


    with x1:

        selected = st.checkbox(

            f"{number} {label}",

            value=(
                st.session_state
                .accessory_qty
                .get(part, 0)
                > 0
            ),

            key=f"acc_{part}",
        )


    with x2:

        if selected:

            qty = st.number_input(

                "Qty",

                min_value=1,

                max_value=999,

                value=int(
                    st.session_state
                    .accessory_qty
                    .get(part, 1)
                ),

                step=1,

                key=f"acc_qty_{part}",
            )


            st.session_state.accessory_qty[
                part
            ] = qty

        else:

            st.session_state.accessory_qty.pop(
                part,
                None
            )


# ============================================================
# BUILD BOM
# ============================================================

bom = build_bom()


# ============================================================
# COST VALUES
# ============================================================

base_cost, optional_cost, pdu_cost, total_cost = (
    cost_summary(bom)
)


margin_pct = (
    st.session_state.margin_pct
)

freight = (
    st.session_state.freight
)

installation = (
    st.session_state.installation
)

warranty_pct = (
    st.session_state.warranty_pct
)


# ============================================================
# 4. COST SUMMARY
# INTERNAL ONLY
# ============================================================

if is_internal:

    ribbon(
        "4. Cost Summary"
    )


    a, b, c, d = st.columns(4)


    with a:

        price_box(
            "Base Cost",
            base_cost
        )


    with b:

        price_box(
            "Optional Cost",
            optional_cost
        )


    with c:

        price_box(
            "PDU Cost",
            pdu_cost
        )


    with d:

        price_box(
            "Total Cost",
            total_cost
        )


# ============================================================
# 5. COST TO PRICE CONVERSION
# INTERNAL ONLY
# ============================================================

if is_internal:

    ribbon(
        "5. Cost to Price Conversion"
    )


    p1, p2, p3, p4 = st.columns(
        4
    )


    with p1:

        st.session_state.margin_pct = (
            st.number_input(

                "Margin (%)",

                0.0,

                99.0,

                st.session_state.margin_pct,

                0.5
            )
        )


    with p2:

        st.session_state.freight = (
            st.number_input(

                "Freight",

                0.0,

                value=
                st.session_state.freight,

                step=500.0
            )
        )


    with p3:

        st.session_state.installation = (
            st.number_input(

                "Installation",

                0.0,

                value=
                st.session_state.installation,

                step=500.0
            )
        )


    with p4:

        st.session_state.warranty_pct = (
            st.number_input(

                "Warranty (%)",

                0.0,

                100.0,

                st.session_state.warranty_pct,

                0.5
            )
        )


    margin_pct = (
        st.session_state.margin_pct
    )

    freight = (
        st.session_state.freight
    )

    installation = (
        st.session_state.installation
    )

    warranty_pct = (
        st.session_state.warranty_pct
    )


    margin_price = (

        total_cost
        /
        (1 - margin_pct / 100)

        if margin_pct < 100

        else 0
    )


    final_selling_price = (

        margin_price
        + freight
        + installation
    )


    warranty_amount = (

        margin_price
        * warranty_pct
        / 100
    )


    q1, q2, q3, q4 = st.columns(
        4
    )


    with q1:

        price_box(
            "Margin Price",
            margin_price
        )


    with q2:

        price_box(
            "After Freight",
            margin_price + freight
        )


    with q3:

        price_box(
            "Final Selling Price",
            final_selling_price,
            final=True
        )


    with q4:

        price_box(
            "Warranty Amount",
            warranty_amount
        )


else:

    margin_price = total_cost

    final_selling_price = total_cost

    warranty_amount = 0.0


# ============================================================
# 6. FINAL BOQ
# ============================================================

ribbon(
    "6. Final BOQ"
)


bom_with_price, margin_price, final_selling_price = (
    add_selling_prices(

        bom,

        total_cost,

        margin_pct,

        freight,

        installation
    )
)


if not bom_with_price.empty:

    # --------------------------------------------------------
    # PREVIOUS BOQ STRUCTURE / NUMBERING
    # --------------------------------------------------------

    structure = bom_with_price[
        [
            "S.No.",
            "Part Code",
            "Description",
            "Quantity",
            "UOM"
        ]
    ].copy()


    MAIN_MDC_PART = "801029209"


    selected_config_components = (
        selected_components()
    )


    cooling_part_codes = set()


    if not selected_config_components.empty:

        cooling_rows = (
            selected_config_components.tail(3)
        )


        cooling_part_codes = set(

            cooling_rows["Part Code"]
            .dropna()
            .astype(str)
            .str.strip()
        )


    # --------------------------------------------------------
    # NUMBERING
    # --------------------------------------------------------

    new_serial = []

    main_mdc_found = False

    mdc_sub_no = 0

    cooling_started = False

    cooling_sub_no = 0

    accessory_no = 3


    for row_index, row in structure.iterrows():

        part_code = clean_part(
            row["Part Code"]
        )

        description = str(
            row["Description"]
        ).strip()


        component_type = str(

            bom_with_price.loc[
                row.name,
                "Component Type"
            ]

        ).strip()


        # ----------------------------------------------------
        # MAIN MDC TITLE
        # ----------------------------------------------------

        if (

            not main_mdc_found

            and
            "SINGLE RACK MDC"
            in description.upper()

        ):

            new_serial.append("")

            main_mdc_found = True

            continue


        # ----------------------------------------------------
        # MAIN MDC
        # ----------------------------------------------------

        if part_code == MAIN_MDC_PART:

            new_serial.append("1")

            continue


        # ----------------------------------------------------
        # COOLING UNIT
        # ----------------------------------------------------

        if part_code in cooling_part_codes:

            cooling_started = True

            cooling_sub_no += 1

            new_serial.append(
                f"2.{cooling_sub_no}"
            )

            continue


        # ----------------------------------------------------
        # OPTIONAL ACCESSORIES
        # ----------------------------------------------------

        if (
            component_type
            == "Optional Accessory"
        ):

            new_serial.append(
                str(accessory_no)
            )

            accessory_no += 1

            continue


        # ----------------------------------------------------
        # PDU
        # ----------------------------------------------------

        if component_type == "PDU":

            new_serial.append(
                str(accessory_no)
            )

            accessory_no += 1

            continue


        # ----------------------------------------------------
        # MDC COMPONENTS
        # ----------------------------------------------------

        if not cooling_started:

            mdc_sub_no += 1

            new_serial.append(
                f"1.{mdc_sub_no}"
            )

            continue


        # ----------------------------------------------------
        # FALLBACK
        # ----------------------------------------------------

        new_serial.append(
            str(accessory_no)
        )

        accessory_no += 1


    structure[
        "New S.No."
    ] = new_serial


    # ========================================================
    # BOQ HTML
    # ========================================================

    html = """

    <table class="final-boq-table">

    <thead>

    <tr>

        <th style="width:7%">
            S.No.
        </th>

        <th style="width:14%">
            Part Code
        </th>

        <th style="width:54%">
            Description
        </th>

        <th style="width:9%">
            Qty
        </th>

        <th style="width:7%">
            UOM
        </th>

        <th style="width:9%">
            Unit Price
        </th>

        <th style="width:10%">
            Total Price
        </th>

    </tr>

    </thead>

    <tbody>

    """


    cooling_heading_added = False

    accessories_heading_added = False

    pdu_heading_added = False


    for row_index, row in structure.iterrows():

        part_code = clean_part(
            row["Part Code"]
        )

        description = str(
            row["Description"]
        ).strip()

        quantity = str(
            row["Quantity"]
        ).strip()

        uom = str(
            row["UOM"]
        ).strip()

        serial_no = str(
            row["New S.No."]
        ).strip()


        component_type = str(

            bom_with_price.loc[
                row.name,
                "Component Type"
            ]

        ).strip()


        unit_price = (
            bom_with_price.loc[
                row.name,
                "Unit Price"
            ]
        )


        total_price = (
            bom_with_price.loc[
                row.name,
                "Total Price"
            ]
        )


        unit_price_text = (

            money(unit_price)

            if pd.notna(unit_price)

            else "N/A"
        )


        total_price_text = (

            money(total_price)

            if pd.notna(total_price)

            else "N/A"
        )


        # ----------------------------------------------------
        # MAIN MDC TITLE
        # ----------------------------------------------------

        if (

            serial_no == ""

            and
            "SINGLE RACK MDC"
            in description.upper()

        ):

            html += f"""

            <tr class="main-mdc-row">

                <td colspan="7">

                    {description}

                </td>

            </tr>

            """

            continue


        # ----------------------------------------------------
        # COOLING UNIT HEADING
        # ----------------------------------------------------

        if (

            part_code in cooling_part_codes

            and
            not cooling_heading_added

        ):

            html += """

            <tr class="section-row">

                <td colspan="7">

                    COOLING UNIT

                </td>

            </tr>

            """

            cooling_heading_added = True


        # ----------------------------------------------------
        # OTHER ACCESSORIES HEADING
        # ----------------------------------------------------

        if (

            component_type
            == "Optional Accessory"

            and
            not accessories_heading_added

        ):

            html += """

            <tr class="section-row">

                <td colspan="7">

                    OTHER ACCESSORIES

                </td>

            </tr>

            """

            accessories_heading_added = True


        # ----------------------------------------------------
        # PDU HEADING
        # ----------------------------------------------------

        if (

            component_type
            == "PDU"

            and
            not pdu_heading_added

        ):

            html += """

            <tr class="section-row">

                <td colspan="7">

                    PDU

                </td>

            </tr>

            """

            pdu_heading_added = True


        # ----------------------------------------------------
        # NORMAL ROW
        # ----------------------------------------------------

        html += f"""

        <tr>

            <td class="center">
                {serial_no}
            </td>

            <td>
                {part_code}
            </td>

            <td>
                {description}
            </td>

            <td class="center">
                {quantity}
            </td>

            <td class="center">
                {uom}
            </td>

            <td>
                {unit_price_text}
            </td>

            <td>
                {total_price_text}
            </td>

        </tr>

        """


    html += """

    </tbody>

    </table>

    """


    # --------------------------------------------------------
    # CONFIGURATION NAME ABOVE TABLE
    # --------------------------------------------------------

    selected_name = CONFIG_NAMES.get(

        st.session_state.configuration,

        st.session_state.configuration
    )


    st.markdown(
        f"""
        <div class="selection-card">

            {selected_name}

        </div>
        """,
        unsafe_allow_html=True,
    )


    st.html(html)


else:

    st.info(
        "No BOQ components available for the selected configuration."
    )


# ============================================================
# EXPORT FUNCTIONS
# ============================================================

def export_excel(
    bom_export,
    internal=False
):

    output = BytesIO()


    selected_name = CONFIG_NAMES.get(

        st.session_state.configuration,

        st.session_state.configuration
    )


    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # HEADER INFORMATION
        # ----------------------------------------------------

        ws_data = [

            [
                "Eaton MDC Solution Configurator"
            ],

            [
                "User Code",
                current_user_code()
            ],

            [
                "Date",
                datetime.now().strftime(
                    "%d-%m-%Y %H:%M"
                )
            ],

            [
                "Customer Name",
                st.session_state.customer_name
            ],

            [
                "Customer Place",
                st.session_state.customer_place
            ],

            [
                "MDC Type",
                st.session_state.mdc_type
            ],

            [
                "Configuration",
                selected_name
            ],

            [],

            [
                "Final BOQ"
            ],
        ]


        meta_df = pd.DataFrame(
            ws_data
        )


        meta_df.to_excel(

            writer,

            index=False,

            header=False,

            sheet_name="MDC BOQ",

            startrow=0
        )


        # ----------------------------------------------------
        # FINAL BOQ
        # ----------------------------------------------------

        sales_cols = [

            "S.No.",

            "Part Code",

            "Description",

            "Quantity",

            "UOM",

            "Unit Price",

            "Total Price"
        ]


        export_df = (
            bom_export[
                sales_cols
            ].copy()
        )


        export_df.to_excel(

            writer,

            index=False,

            sheet_name="MDC BOQ",

            startrow=len(ws_data) + 1
        )


        # ----------------------------------------------------
        # COST SUMMARY
        # ----------------------------------------------------

        start = (

            len(ws_data)
            +
            len(export_df)
            +
            4
        )


        summary = pd.DataFrame(

            [

                [
                    "Final Selling Price",
                    final_selling_price
                ],

                [
                    "Base Cost",
                    (
                        base_cost
                        if internal
                        else "Not shown"
                    )
                ],

                [
                    "Optional Cost",
                    (
                        optional_cost
                        if internal
                        else "Not shown"
                    )
                ],

                [
                    "PDU Cost",
                    (
                        pdu_cost
                        if internal
                        else "Not shown"
                    )
                ],

                [
                    "Total Cost",
                    (
                        total_cost
                        if internal
                        else "Not shown"
                    )
                ],

                [
                    "Margin %",
                    (
                        margin_pct
                        if internal
                        else "Not shown"
                    )
                ],

                [
                    "Freight",
                    (
                        freight
                        if internal
                        else "Not shown"
                    )
                ],

                [
                    "Installation",
                    (
                        installation
                        if internal
                        else "Not shown"
                    )
                ],

                [
                    "Warranty %",
                    (
                        warranty_pct
                        if internal
                        else "Not shown"
                    )
                ],
            ],

            columns=[
                "Item",
                "Value"
            ]
        )


        summary.to_excel(

            writer,

            index=False,

            sheet_name="MDC BOQ",

            startrow=start
        )


    output.seek(0)

    return output.getvalue()


# ============================================================
# PDF EXPORT
# ============================================================

def export_pdf(
    bom_export,
    internal=False
):

    from reportlab.lib import colors

    from reportlab.lib.enums import (
        TA_CENTER,
        TA_LEFT
    )

    from reportlab.lib.pagesizes import (
        A4,
        landscape
    )

    from reportlab.lib.styles import (
        getSampleStyleSheet,
        ParagraphStyle
    )

    from reportlab.lib.units import mm

    from reportlab.platypus import (
        SimpleDocTemplate,
        Table,
        TableStyle,
        Paragraph,
        Spacer
    )


    output = BytesIO()


    selected_name = CONFIG_NAMES.get(

        st.session_state.configuration,

        st.session_state.configuration
    )


    doc = SimpleDocTemplate(

        output,

        pagesize=landscape(A4),

        rightMargin=8 * mm,

        leftMargin=8 * mm,

        topMargin=8 * mm,

        bottomMargin=8 * mm,
    )


    styles = getSampleStyleSheet()


    title_style = ParagraphStyle(

        "MDCTitle",

        parent=styles["Title"],

        fontName="Helvetica-Bold",

        fontSize=16,

        leading=18,

        textColor=colors.HexColor(
            EATON_BLUE
        ),

        alignment=TA_LEFT,

        spaceAfter=3,
    )


    small = ParagraphStyle(

        "Small",

        parent=styles["Normal"],

        fontSize=7.5,

        leading=9,

        textColor=colors.HexColor(
            TEXT
        ),
    )


    small_center = ParagraphStyle(

        "SmallCenter",

        parent=small,

        alignment=TA_CENTER,
    )


    story = []


    # --------------------------------------------------------
    # PDF TITLE
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "Eaton MDC Solution Configurator",
            title_style
        )
    )


    story.append(
        Paragraph(
            "Modular Data Center Solution Configuration & Pricing",
            small
        )
    )


    story.append(
        Spacer(
            1,
            3 * mm
        )
    )


    # --------------------------------------------------------
    # PDF META
    # --------------------------------------------------------

    meta = [

        [

            Paragraph(
                f"<b>User Code:</b> "
                f"{current_user_code()}",
                small
            ),

            Paragraph(
                f"<b>Date:</b> "
                f"{datetime.now().strftime('%d-%m-%Y %H:%M')}",
                small
            ),

            Paragraph(
                f"<b>Customer:</b> "
                f"{st.session_state.customer_name}",
                small
            ),

            Paragraph(
                f"<b>MDC Type:</b> "
                f"{st.session_state.mdc_type}",
                small
            ),
        ],

        [

            Paragraph(
                f"<b>Configuration:</b> "
                f"{selected_name}",
                small
            ),

            Paragraph(
                f"<b>Customer Place:</b> "
                f"{st.session_state.customer_place}",
                small
            ),

            Paragraph(
                f"<b>Final Selling Price:</b> "
                f"{money(final_selling_price)}",
                small
            ),

            Paragraph(
                f"<b>Access:</b> "
                f"{'Internal' if internal else 'Sales'}",
                small
            ),
        ],
    ]


    mt = Table(

        meta,

        colWidths=[

            68 * mm,

            48 * mm,

            68 * mm,

            62 * mm
        ]
    )


    mt.setStyle(
        TableStyle(

            [

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor(
                        "#F1F6FA"
                    )
                ),

                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.HexColor(
                        BORDER
                    )
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.3,
                    colors.HexColor(
                        BORDER
                    )
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
            ]
        )
    )


    story.append(mt)


    story.append(
        Spacer(
            1,
            4 * mm
        )
    )


    story.append(
        Paragraph(
            "Final BOQ",
            styles["Heading2"]
        )
    )


    # --------------------------------------------------------
    # BOQ TABLE
    # --------------------------------------------------------

    data = [

        [

            Paragraph(
                "S.No.",
                small_center
            ),

            Paragraph(
                "Part Code",
                small_center
            ),

            Paragraph(
                "Description",
                small_center
            ),

            Paragraph(
                "Qty",
                small_center
            ),

            Paragraph(
                "UOM",
                small_center
            ),

            Paragraph(
                "Unit Price",
                small_center
            ),

            Paragraph(
                "Total Price",
                small_center
            ),
        ]
    ]


    for _, row in bom_export.iterrows():

        data.append(

            [

                Paragraph(
                    str(row["S.No."]),
                    small_center
                ),

                Paragraph(
                    clean_part(
                        row["Part Code"]
                    ),
                    small
                ),

                Paragraph(
                    str(
                        row["Description"]
                    ),
                    small
                ),

                Paragraph(
                    str(
                        row["Quantity"]
                    ),
                    small_center
                ),

                Paragraph(
                    str(
                        row["UOM"]
                    ),
                    small_center
                ),

                Paragraph(

                    money(
                        row["Unit Price"]
                    )

                    if
                    pd.notna(
                        row["Unit Price"]
                    )

                    else "N/A",

                    small
                ),

                Paragraph(

                    money(
                        row["Total Price"]
                    )

                    if
                    pd.notna(
                        row["Total Price"]
                    )

                    else "N/A",

                    small
                ),
            ]
        )


    t = Table(

        data,

        repeatRows=1,

        colWidths=[

            15 * mm,

            33 * mm,

            105 * mm,

            15 * mm,

            15 * mm,

            35 * mm,

            35 * mm
        ]
    )


    t.setStyle(

        TableStyle(

            [

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        "#F1F6FA"
                    )
                ),

                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        EATON_DARK
                    )
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold"
                ),

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.3,
                    colors.HexColor(
                        BORDER
                    )
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    3
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    3
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    3
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    3
                ),
            ]
        )
    )


    story.append(t)


    story.append(
        Spacer(
            1,
            4 * mm
        )
    )


    # --------------------------------------------------------
    # PRICE SUMMARY
    # --------------------------------------------------------

    summary_data = [

        [

            Paragraph(
                "Final Selling Price",
                small
            ),

            Paragraph(
                money(
                    final_selling_price
                ),
                small
            )
        ]
    ]


    if internal:

        summary_data.extend(

            [

                [

                    Paragraph(
                        "Base Cost",
                        small
                    ),

                    Paragraph(
                        money(base_cost),
                        small
                    )
                ],

                [

                    Paragraph(
                        "Optional Cost",
                        small
                    ),

                    Paragraph(
                        money(optional_cost),
                        small
                    )
                ],

                [

                    Paragraph(
                        "PDU Cost",
                        small
                    ),

                    Paragraph(
                        money(pdu_cost),
                        small
                    )
                ],

                [

                    Paragraph(
                        "Total Cost",
                        small
                    ),

                    Paragraph(
                        money(total_cost),
                        small
                    )
                ],
            ]
        )


    stbl = Table(

        summary_data,

        colWidths=[

            45 * mm,

            45 * mm
        ]
    )


    stbl.setStyle(

        TableStyle(

            [

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        EATON_LIGHT
                    )
                ),

                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.HexColor(
                        BORDER
                    )
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.3,
                    colors.HexColor(
                        BORDER
                    )
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold"
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
            ]
        )
    )


    story.append(stbl)


    doc.build(story)


    output.seek(0)

    return output.getvalue()


# ============================================================
# 7 FINAL SELLING PRICE + DOWNLOADS
# ============================================================

ribbon(
    "Final Selling Price & Downloads"
)


# ------------------------------------------------------------
# Customer mandatory check
# ------------------------------------------------------------

customer_valid = bool(

    st.session_state.customer_name
    .strip()
)


if not customer_valid:

    st.warning(

        "Customer Name is mandatory. "
        "Enter the customer name to enable "
        "Excel and PDF downloads."
    )


# ------------------------------------------------------------
# Generate files
# ------------------------------------------------------------

if not bom.empty:

    final_excel = export_excel(

        bom_with_price,

        internal=is_internal
    )


    final_pdf = export_pdf(

        bom_with_price,

        internal=is_internal
    )


    d1, d2 = st.columns(
        [1.15, 1.0]
    )


    # --------------------------------------------------------
    # Final selling price LEFT
    # --------------------------------------------------------

    with d1:

        price_box(

            "Final Selling Price",

            final_selling_price,

            final=True
        )


    # --------------------------------------------------------
    # Downloads RIGHT
    # --------------------------------------------------------

    with d2:

        b1, b2 = st.columns(2)


        with b1:

            st.download_button(

                "Download Excel",

                data=final_excel,

                file_name=
                "MDC_Final_BOQ.xlsx",

                mime=
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                use_container_width=True,

                disabled=
                not customer_valid,

                on_click=
                record_download,

                args=(
                    next_code,
                    "Excel"
                ),
            )


        with b2:

            st.download_button(

                "Download PDF",

                data=final_pdf,

                file_name=
                "MDC_Final_BOQ.pdf",

                mime=
                "application/pdf",

                use_container_width=True,

                disabled=
                not customer_valid,

                on_click=
                record_download,

                args=(
                    next_code,
                    "PDF"
                ),
            )


    if not is_internal:

        st.markdown(

            """
            <div class="compact-note">

                Sales view contains selling prices only.
                Internal cost values are not displayed.

            </div>
            """,

            unsafe_allow_html=True
        )


else:

    st.info(
        "No BOQ available."
    )


# ============================================================
# 8 CONFIGURATION HISTORY
# INTERNAL ONLY
# ============================================================

if is_internal:

    ribbon(
        "Configuration History"
    )


    conn = sqlite3.connect(
        TRACKING_DB
    )


    history_df = pd.read_sql_query(

        """
        SELECT

            user_code AS "User Code",

            downloaded_at AS "Date",

            customer_name AS "Customer Name"

        FROM download_history

        ORDER BY id DESC
        """,

        conn
    )


    conn.close()


    if not history_df.empty:

        st.dataframe(

            history_df,

            use_container_width=True,

            hide_index=True
        )

    else:

        st.info(
            "No downloaded configurations available yet."
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()


st.markdown(

    """
    <div class="compact-note">

        MDC Solution V1 |
        Single Rack data loaded from the supplied BOQ |
        Multirack configurations remain based on the available master data.

    </div>
    """,

    unsafe_allow_html=True
)
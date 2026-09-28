import os
import re
import sqlite3
from io import BytesIO
from datetime import datetime

import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

import streamlit as st


# ============================================================
# EATON MDC SOLUTION CONFIGURATOR
# Data sources:
#   Single Rack  -> MDC_Master_V1.xlsx
#   Multirack    -> MULTIRACK BOQ.xlsx
#                  sheet: Multi Rack config-1
# IMPORTANT:
#   Both workbooks are expected in the same folder as this app.py.
# ============================================================

st.set_page_config(
    page_title="Eaton MDC Solution Configurator",
    page_icon="🏢",
    layout="wide",
)

# Compact desktop UI. This block changes presentation only; it does not
# change the configuration, BOM, pricing, download, or save logic.
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 1500px;
    }
    div[data-testid="stVerticalBlock"] {
        gap: 0.45rem;
    }
    div[data-testid="stHorizontalBlock"] {
        gap: 0.8rem;
    }
    label, .stMarkdown p, .stCaption {
        font-size: 12px !important;
    }
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"],
    div[data-testid="stRadio"] label {
        font-size: 13px !important;
    }
    div[data-testid="stTextInput"],
    div[data-testid="stNumberInput"],
    div[data-testid="stSelectbox"],
    div[data-testid="stRadio"] {
        margin-bottom: 0 !important;
    }
    button[kind] {
        min-height: 34px !important;
        font-size: 13px !important;
    }
    div[data-testid="stDownloadButton"] button {
        min-height: 36px !important;
    }
    hr { margin: 0.5rem 0 !important; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_FILE = os.path.join(BASE_DIR, "MDC_Master_V1.xlsx")
MULTIRACK_FILE = os.path.join(BASE_DIR, "MULTIRACK BOQ.xlsx")
MULTIRACK_SHEET = "Multi Rack config-1"
TRACKING_DB = os.path.join(BASE_DIR, "MDC_Tracking.db")
DEMO_INTERNAL_PASSWORD = "MDC@123"


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def numeric(value):
    """Return a float for valid Excel numbers; otherwise NaN."""
    try:
        if pd.isna(value):
            return float("nan")
        if isinstance(value, str):
            value = value.strip().replace(",", "")
            if value in ("", "#N/A", "N/A", "NA", "nan", "None"):
                return float("nan")
        return float(value)
    except Exception:
        return float("nan")


def money(value):
    """Format a numeric price; keep missing Excel values completely blank."""
    try:
        if pd.isna(value):
            return ""
        return f"₹ {float(value):,.2f}"
    except Exception:
        return ""


def internal_password():
    try:
        return st.secrets["MDC_INTERNAL_PASSWORD"]
    except Exception:
        return DEMO_INTERNAL_PASSWORD


# ============================================================
# DATABASE
# ============================================================

def init_tracking_db():
    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    cur.execute("""
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
            warranty_amount REAL,
            user_code TEXT
        )
    """)

    # Upgrade an older database created by the previous app.
    cols = [r[1] for r in cur.execute(
        "PRAGMA table_info(configurations)"
    ).fetchall()]

    if "user_code" not in cols:
        cur.execute("ALTER TABLE configurations ADD COLUMN user_code TEXT")

    cur.execute("""
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
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS download_counter (
            id INTEGER PRIMARY KEY,
            download_count INTEGER NOT NULL DEFAULT 0
        )
    """)

    cur.execute("""
        INSERT OR IGNORE INTO download_counter (id, download_count)
        VALUES (1, 0)
    """)

    # ========================================================
    # USER SESSION COUNTER
    # Counts a new Streamlit app session only once.
    # This is intentionally separate from download_counter.
    # ========================================================
    cur.execute("""
        CREATE TABLE IF NOT EXISTS session_counter (
            id INTEGER PRIMARY KEY,
            user_count INTEGER NOT NULL DEFAULT 0
        )
    """)

    cur.execute("""
        INSERT OR IGNORE INTO session_counter (id, user_count)
        VALUES (1, 0)
    """)

    # ========================================================
    # USER DOWNLOAD HISTORY
    # ========================================================
    cur.execute("""
        CREATE TABLE IF NOT EXISTS user_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_code TEXT,
            customer_name TEXT,
            created_at TEXT
        )
    """)

    conn.commit()
    conn.close()


def increment_download_count():
    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    cur.execute("""
        UPDATE download_counter
        SET download_count = download_count + 1
        WHERE id = 1
    """)

    count = cur.execute("""
        SELECT download_count
        FROM download_counter
        WHERE id = 1
    """).fetchone()[0]

    conn.commit()
    conn.close()
    return count


def increment_user_count():
    """Increment the persistent count for a new Streamlit user session."""
    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    cur.execute("""
        UPDATE session_counter
        SET user_count = user_count + 1
        WHERE id = 1
    """)

    count = cur.execute("""
        SELECT user_count
        FROM session_counter
        WHERE id = 1
    """).fetchone()[0]

    conn.commit()
    conn.close()
    return count


def generate_configuration_id():
    today = datetime.now().strftime("%Y%m%d")

    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    count = cur.execute("""
        SELECT COUNT(*)
        FROM configurations
        WHERE configuration_id LIKE ?
    """, (f"MDC-{today}-%",)).fetchone()[0] + 1

    conn.close()
    return f"MDC-{today}-{count:04d}"


def save_configuration(
    configuration_id,
    bom,
    base_cost,
    optional_cost,
    pdu_cost,
    total_cost,
    margin_pct,
    freight,
    installation,
    warranty_pct,
    margin_price,
    final_selling_price,
    warranty_amount,
):
    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cur.execute("""
        INSERT OR REPLACE INTO configurations (
            configuration_id,
            created_at,
            customer_name,
            customer_place,
            problem,
            solution,
            mdc_type,
            configuration,
            base_cost,
            optional_cost,
            pdu_cost,
            total_cost,
            margin_pct,
            freight,
            installation,
            warranty_pct,
            margin_price,
            final_selling_price,
            warranty_amount,
            user_code
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        configuration_id,
        created_at,
        st.session_state.customer_name,
        st.session_state.customer_place,
        st.session_state.problem,
        st.session_state.solution,
        st.session_state.mdc_type,
        st.session_state.configuration,
        float(base_cost),
        float(optional_cost),
        float(pdu_cost),
        float(total_cost),
        float(margin_pct),
        float(freight),
        float(installation),
        float(warranty_pct),
        float(margin_price),
        float(final_selling_price),
        float(warranty_amount),
        st.session_state.user_code,
    ))

    cur.execute(
        "DELETE FROM configuration_items WHERE configuration_id = ?",
        (configuration_id,),
    )

    if bom is not None and not bom.empty:
        for _, row in bom.iterrows():
            cur.execute("""
                INSERT INTO configuration_items (
                    configuration_id,
                    component_type,
                    part_code,
                    description,
                    quantity,
                    uom,
                    unit_cost,
                    total_cost,
                    unit_price,
                    total_price
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                configuration_id,
                clean_text(row.get("Component Type")),
                clean_text(row.get("Part Code")),
                clean_text(row.get("Description")),
                float(numeric(row.get("Quantity", 0))) if pd.notna(row.get("Quantity")) else 0,
                clean_text(row.get("UOM")),
                float(numeric(row.get("Unit Cost"))) if pd.notna(row.get("Unit Cost")) else 0,
                float(numeric(row.get("Total Cost"))) if pd.notna(row.get("Total Cost")) else 0,
                float(numeric(row.get("Unit Price"))) if pd.notna(row.get("Unit Price")) else None,
                float(numeric(row.get("Total Price"))) if pd.notna(row.get("Total Price")) else None,
            ))

    conn.commit()
    conn.close()


init_tracking_db()


# ============================================================
# LOAD THE ONE-SHEET EXCEL MASTER
# EVERYTHING USED BY THE UI IS READ FROM MDC_Master_V1.xlsx
# ============================================================

def file_mtime(path):
    """Return the file modification time so Streamlit reloads changed Excel files."""
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0


@st.cache_data

def load_master(master_mtime=0.0, multirack_mtime=0.0):
    """
    Load the existing Single Rack master exactly as before and, separately,
    load the Multirack workbook when it is available.

    Multirack source:
        MULTIRACK BOQ.xlsx
        sheet: Multi Rack config-1

    The Multirack workbook is intentionally treated as a non-priced BOQ:
        S.No. | Description | Qty | UOM

    The parser is tolerant of the workbook's layout. It detects:
      - Configuration 1 ... Configuration 9
      - B / C group headings
      - S.No / Description / Qty / UOM header rows
    """
    if not os.path.exists(MASTER_FILE):
        raise FileNotFoundError(
            f"MDC_Master_V1.xlsx was not found in: {BASE_DIR}"
        )

    sheet = pd.read_excel(
        MASTER_FILE,
        sheet_name=0,
        header=None,
        engine="openpyxl",
    )

    def get(row, col):
        if col >= len(row):
            return ""
        return clean_text(row.iloc[col])

    def is_part_header(value):
        return get(pd.Series([value]), 0).upper() in {
            "PART NUMBER", "PART NO", "PART CODE"
        }

    # --------------------------------------------------------
    # SINGLE-RACK CONFIGURATIONS
    # --------------------------------------------------------
    block_info = {
        "Configuration 1": {
            "start": 0, "end": 25,
            "part": 0, "desc": 1, "qty": 2,
            "uom": 3, "price": 4,
        },
        "Configuration 3": {
            "start": 0, "end": 25,
            "part": 5, "desc": 6, "qty": 7,
            "uom": 8, "price": 9,
        },
        "Configuration 2": {
            "start": 25, "end": 50,
            "part": 0, "desc": 1, "qty": 2,
            "uom": 3, "price": 4,
        },
        "Configuration 4": {
            "start": 25, "end": 50,
            "part": 5, "desc": 6, "qty": 7,
            "uom": 8, "price": 9,
        },
    }

    config_rows = []
    component_rows = []

    for config_name, info in block_info.items():
        block = sheet.iloc[info["start"]:info["end"]]

        title_col = info["part"]
        solution_title = get(block.iloc[0], title_col)
        if not solution_title:
            solution_title = config_name

        solution_title = " ".join(solution_title.split())
        solution_title = solution_title.replace("SOLUTION 1", "")
        solution_title = solution_title.replace("SOLUTION 2", "")
        solution_title = solution_title.replace("SOLUTION 3", "")
        solution_title = solution_title.replace("SOLUTION 4", "")
        solution_title = " ".join(solution_title.split()).strip(" -")

        for local_index, (_, row) in enumerate(block.iterrows()):
            if local_index in (0, 1):
                continue

            part = get(row, info["part"])
            desc = get(row, info["desc"])
            qty = numeric(row.iloc[info["qty"]])
            uom = get(row, info["uom"])
            price = numeric(row.iloc[info["price"]])

            if not part and not desc:
                continue

            if part.upper() in {"PART NUMBER", "PART NO", "PART CODE"}:
                continue

            if part.upper() == "CTO3M002":
                continue

            component_rows.append({
                "MDC Type": "Single Rack",
                "Configuration": config_name,
                "Configuration Title": solution_title,
                "Part Code": part,
                "Description": desc,
                "Quantity": 0.0 if pd.isna(qty) else float(qty),
                "UOM": uom if uom else "EA",
                "Unit Cost": price,
            })

        cfg_components = pd.DataFrame([
            r for r in component_rows
            if r["Configuration"] == config_name
        ])

        if cfg_components.empty:
            base_cost = 0.0
        else:
            base_cost = float(
                pd.to_numeric(
                    cfg_components["Unit Cost"], errors="coerce"
                ).fillna(0).mul(
                    pd.to_numeric(
                        cfg_components["Quantity"], errors="coerce"
                    ).fillna(0)
                ).sum()
            )

        config_rows.append({
            "MDC Type": "Single Rack",
            "Configuration": config_name,
            "Configuration Title": solution_title,
            "Base Cost": base_cost,
        })

    configs = pd.DataFrame(config_rows)
    components = pd.DataFrame(component_rows)

    # --------------------------------------------------------
    # MULTIRACK BOQ
    # --------------------------------------------------------
    multirack_configs, multirack_components = load_multirack_boq()

    if multirack_configs.empty:
        # Keep the existing app usable even before the new GitHub Excel
        # file is uploaded. The nine configuration selectors are still
        # created and the UI will show a clear data-source message.
        multirack_configs = pd.DataFrame([
            {
                "MDC Type": "Multirack",
                "Configuration": f"Configuration {n}",
                "Configuration Title": f"Multirack Configuration {n}",
                "Base Cost": 0.0,
            }
            for n in range(1, 10)
        ])

    configs = pd.concat(
        [configs, multirack_configs],
        ignore_index=True,
    )

    # --------------------------------------------------------
    # OTHER OPTIONAL ITEMS
    # --------------------------------------------------------
    accessories = []
    optional_start = None
    optional_end = None

    for i in range(len(sheet)):
        text = " ".join(get(sheet.iloc[i], 0).split()).upper()
        if "OTHER OPTIONAL ITEMS" in text:
            optional_start = i + 1
            continue
        if optional_start is not None and "SINGLE PHASE PDU" in text:
            optional_end = i
            break

    if optional_start is not None:
        if optional_end is None:
            optional_end = len(sheet)

        for _, row in sheet.iloc[optional_start:optional_end].iterrows():
            part = get(row, 0)
            desc = get(row, 1)
            qty = numeric(row.iloc[2])
            uom = get(row, 3)
            price = numeric(row.iloc[4])

            if not part or not desc:
                continue

            if part.upper() in {"PART NUMBER", "PART NO", "PART CODE"}:
                continue

            accessories.append({
                "Part Code": part,
                "Description": desc,
                "Default Quantity": (
                    1.0 if pd.isna(qty) or qty <= 0 else float(qty)
                ),
                "UOM": uom if uom else "EA",
                "Unit Cost": price,
            })

    accessories = pd.DataFrame(accessories)

    # --------------------------------------------------------
    # SINGLE PHASE PDU'S
    # --------------------------------------------------------
    pdus = []
    pdu_start = None

    for i in range(len(sheet)):
        text = " ".join(get(sheet.iloc[i], 0).split()).upper()
        if "SINGLE PHASE PDU" in text:
            pdu_start = i + 1
            break

    current_pdu_type = ""

    if pdu_start is not None:
        for _, row in sheet.iloc[pdu_start:].iterrows():
            part = get(row, 0)
            desc = get(row, 1)

            if not part and not desc:
                continue

            if part.upper() in {"PART NUMBER", "PART NO", "PART CODE"}:
                continue

            c13 = numeric(row.iloc[2])
            c19 = numeric(row.iloc[3])
            excel_type = get(row, 4)
            price = numeric(row.iloc[5])

            if excel_type:
                current_pdu_type = excel_type.upper()

            if not part or not desc or not current_pdu_type:
                continue

            pdus.append({
                "Part Code": part,
                "Description": desc,
                "C13": 0.0 if pd.isna(c13) else float(c13),
                "C19": 0.0 if pd.isna(c19) else float(c19),
                "Type": current_pdu_type,
                "UOM": "EA",
                "Unit Cost": price,
            })

    pdus = pd.DataFrame(pdus)

    for df in (configs, components, accessories, pdus, multirack_components):
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].fillna("").astype(str).str.strip()

    return configs, components, accessories, pdus, multirack_components


def load_multirack_boq():
    """Read MULTIRACK BOQ.xlsx directly at runtime; the workbook is the source of truth."""
    empty_configs = pd.DataFrame(columns=[
        "MDC Type", "Configuration", "Configuration Title", "Base Cost"
    ])
    empty_components = pd.DataFrame(columns=[
        "MDC Type", "Configuration", "Configuration Title", "Group",
        "Group Label", "S.No.", "Description", "Quantity", "UOM",
        "Part Code", "Unit Cost", "Selection Key"
    ])

    if not os.path.exists(MULTIRACK_FILE):
        return empty_configs, empty_components

    try:
        xl = pd.ExcelFile(MULTIRACK_FILE, engine="openpyxl")
        if MULTIRACK_SHEET in xl.sheet_names:
            sheet_name = MULTIRACK_SHEET
        else:
            candidates = [
                s for s in xl.sheet_names
                if "multirack" in re.sub(r"[^a-z0-9]", "", s.lower())
                or "multi rack" in s.lower()
            ]
            sheet_name = candidates[0] if candidates else xl.sheet_names[0]
        raw = pd.read_excel(
            MULTIRACK_FILE, sheet_name=sheet_name, header=None, engine="openpyxl"
        )
    except Exception:
        return empty_configs, empty_components

    if raw.empty:
        return empty_configs, empty_components

    def cell(row, col):
        if col is None or col >= len(row):
            return ""
        value = row.iloc[col]
        return "" if pd.isna(value) else str(value).strip()

    def row_values(row):
        return [cell(row, c) for c in range(len(raw.columns))]

    def nonempty_values(row):
        return [v for v in row_values(row) if v]

    def norm(value):
        return re.sub(r"[^A-Z0-9]+", "", clean_text(value).upper())

    solution_re = re.compile(
        r"^\s*(?:SOLUTION|SOLN|CONFIGURATION|CONFIG)\s*[-_:.)#]*\s*(\d{1,2})\s*$",
        re.IGNORECASE,
    )
    group_re = re.compile(
        r"^\s*([ABC])\s*(?:[-:.)]\s*(.*))?$",
        re.IGNORECASE,
    )

    solution_starts = []
    for idx in range(len(raw)):
        vals = nonempty_values(raw.iloc[idx])
        if len(vals) == 1:
            match = solution_re.match(vals[0])
            if match:
                number = int(match.group(1))
                if 1 <= number <= 9:
                    solution_starts.append((idx, number, vals[0]))

    unique = {}
    for idx, number, title in solution_starts:
        unique.setdefault(number, (idx, title))
    solution_starts = sorted(
        [(idx, number, title) for number, (idx, title) in unique.items()],
        key=lambda x: x[0],
    )

    if not solution_starts:
        return empty_configs, empty_components

    def find_header(start, end):
        aliases = {
            "sno": {"SNO", "SERIALNO", "SERIALNUMBER", "SRNO"},
            "description": {"DESCRIPTION", "DESC", "ITEMDESCRIPTION"},
            "qty": {"QTY", "QUANTITY", "COUNT"},
            "uom": {"UOM", "UNITOFMEASURE", "UNIT"},
            "part": {"PARTCODE", "PARTNUMBER", "PARTNO", "PART"},
            "cost": {"COST", "UNITCOST", "UNITPRICE", "PRICE", "SELLINGPRICE"},
        }
        for ridx in range(start, end):
            mapped = {}
            for col, value in enumerate(row_values(raw.iloc[ridx])):
                n = norm(value)
                for key, candidates in aliases.items():
                    if n in candidates and key not in mapped:
                        mapped[key] = col
            if all(k in mapped for k in ("sno", "description", "qty", "uom")):
                return ridx, mapped
        return None, {"sno": 0, "description": 1, "qty": 2, "uom": 3}

    config_rows = []
    component_rows = []

    for pos, (start_row, number, solution_title_raw) in enumerate(solution_starts):
        end_row = solution_starts[pos + 1][0] if pos + 1 < len(solution_starts) else len(raw)
        config_name = f"Configuration {number}"
        solution_title = " ".join(clean_text(solution_title_raw).split()) or f"Solution {number}"
        header_row, column_map = find_header(start_row, end_row)
        data_start = header_row + 1 if header_row is not None else start_row + 1

        current_group = "A"
        current_group_label = "A"

        for ridx in range(data_start, end_row):
            row = raw.iloc[ridx]
            meaningful = nonempty_values(row)
            if not meaningful:
                continue

            # The Multirack workbook supports both layouts:
            #   1) B/C appear as standalone group-heading rows.
            #   2) B/C are written directly in the S.No. column of each
            #      optional component row.
            #
            # Keep the existing group-heading format for compatibility.
            if len(meaningful) == 1:
                group_match = group_re.match(meaningful[0])
                if group_match:
                    current_group = group_match.group(1).upper()
                    suffix = clean_text(group_match.group(2))
                    current_group_label = (
                        f"{current_group} - {suffix}" if suffix else current_group
                    )
                    continue

            sno = cell(row, column_map.get("sno"))
            description = cell(row, column_map.get("description"))
            qty_col = column_map.get("qty")
            quantity = numeric(cell(row, qty_col))
            uom = cell(row, column_map.get("uom"))
            part_code = cell(row, column_map.get("part"))
            cost_col = column_map.get("cost")
            unit_cost = numeric(cell(row, cost_col)) if cost_col is not None else float("nan")

            if norm(sno) in {"SNO", "SERIALNO", "SERIALNUMBER", "SRNO"}:
                continue
            if norm(description) in {"DESCRIPTION", "DESC", "ITEMDESCRIPTION"}:
                continue
            if not sno and not description and not part_code:
                continue

            # In the new workbook, B/C in S.No. identify the optional
            # group for that specific row. Normal serial numbers remain
            # mandatory (group A).
            sno_group = norm(sno)
            if sno_group in {"A", "B", "C"}:
                row_group = sno_group
                row_group_label = row_group
                current_group = row_group
                current_group_label = row_group_label
            else:
                row_group = current_group
                row_group_label = current_group_label

            qty_value = float(quantity) if pd.notna(quantity) else 0.0
            selection_key = f"MR-{number}-{row_group}-{ridx}"

            component_rows.append({
                "MDC Type": "Multirack",
                "Configuration": config_name,
                "Configuration Title": solution_title,
                "Group": row_group,
                "Group Label": row_group_label,
                "S.No.": sno,
                "Description": description,
                "Quantity": qty_value,
                "UOM": uom or "EA",
                "Part Code": part_code,
                "Unit Cost": unit_cost,
                "Selection Key": selection_key,
            })

        config_rows.append({
            "MDC Type": "Multirack",
            "Configuration": config_name,
            "Configuration Title": solution_title,
            "Base Cost": 0.0,
        })

    existing = {r["Configuration"] for r in config_rows}
    for number in range(1, 10):
        name = f"Configuration {number}"
        if name not in existing:
            config_rows.append({
                "MDC Type": "Multirack",
                "Configuration": name,
                "Configuration Title": f"Solution {number}",
                "Base Cost": 0.0,
            })

    configs = pd.DataFrame(config_rows)
    configs["_sort"] = configs["Configuration"].str.extract(r"(\d+)")[0].astype(int)
    configs = configs.sort_values("_sort").drop(columns="_sort").reset_index(drop=True)

    components = pd.DataFrame(component_rows)
    if components.empty:
        components = empty_components.copy()

    return configs, components


try:
    configs_df, components_df, accessories_df, pdus_df, multirack_components_df = load_master(
    master_mtime=file_mtime(MASTER_FILE),
    multirack_mtime=file_mtime(MULTIRACK_FILE),
)
except Exception as exc:
    st.error("Unable to load MDC_Master_V1.xlsx.")
    st.exception(exc)
    st.stop()


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
    "margin_pct": 20.0,
    "freight": 100000.0,
    "installation": 150000.0,
    "warranty_pct": 10.0,
    "configuration_id": None,
    "configuration_saved": False,
    "user_code": "—",
    "user_count": 0,
    "session_counted": False,
    "multirack_selected": {},
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

if st.session_state.configuration_id is None:
    st.session_state.configuration_id = generate_configuration_id()

# ============================================================
# NEW USER SESSION COUNT
# Count this Streamlit session exactly once.
#
# Flow:
#   New session  -> User Count +1
#   Same session -> Nothing
#   Download     -> User Count unchanged
# ============================================================
if not st.session_state.session_counted:
    st.session_state.user_count = increment_user_count()
    st.session_state.session_counted = True


# ============================================================
# UI HELPERS
# ============================================================

def section_header(text):
    st.html(f"""
    <div style="
        background:#003B71;
        color:white;
        padding:7px 12px;
        border-radius:5px;
        margin:10px 0 8px 0;
        font-size:14px;
        font-weight:700;
        letter-spacing:.2px;
    ">
        {text}
    </div>
    """)


def price_box(label, value):
    st.markdown(
        f"""
        <div style="padding:2px 0 4px 0;">
            <div style="font-size:12px;color:#475569;font-weight:600;margin-bottom:2px;">
                {label}
            </div>
            <div style="font-size:18px;font-weight:700;color:#003B71;white-space:nowrap;">
                {money(value)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# USER CODE
# ============================================================

def generate_user_code():
    """Generate the user code entirely from the current Excel-driven selections."""
    codes = []

    config_text = clean_text(st.session_state.configuration)
    if config_text:
        number = config_text.split()[-1]
        codes.append(f"C{number}")

    # Fire suppression is detected from Excel descriptions.
    fire_selected = []
    for part, qty in st.session_state.accessory_qty.items():
        if numeric(qty) <= 0:
            continue
        match = accessories_df[
            accessories_df["Part Code"].astype(str).str.strip() == str(part).strip()
        ]
        if not match.empty:
            desc = clean_text(match.iloc[0]["Description"]).upper()
            if "FIRE" in desc:
                fire_selected.append(desc)

    if any("EXTERNAL" in x for x in fire_selected):
        codes.append("F-EXT")
    elif fire_selected:
        codes.append("F-INT")

    # Camera is detected from Excel descriptions, not hardcoded part numbers.
    camera_selected = False
    for part, qty in st.session_state.accessory_qty.items():
        if numeric(qty) <= 0:
            continue
        match = accessories_df[
            accessories_df["Part Code"].astype(str).str.strip() == str(part).strip()
        ]
        if not match.empty and "CAMERA" in clean_text(match.iloc[0]["Description"]).upper():
            camera_selected = True
            break

    if camera_selected:
        codes.append("CAM")

    # Other accessory codes are detected by their Excel descriptions.
    accessory_code_map = [
        ("KEYBOARD", "KT"),
        ("CABLE MANAGER", "CM"),
        ("TOP CABLE TRAY", "TCT"),
        ("BRUSH PANEL", "BP"),
    ]

    for keyword, code in accessory_code_map:
        selected = False
        for part, qty in st.session_state.accessory_qty.items():
            if numeric(qty) <= 0:
                continue
            match = accessories_df[
                accessories_df["Part Code"].astype(str).str.strip() == str(part).strip()
            ]
            if not match.empty and keyword in clean_text(match.iloc[0]["Description"]).upper():
                selected = True
                break
        if selected:
            codes.append(code)

    # PDU code is obtained separately from Excel TYPE for every
    # selected PDU component. This avoids one common PDU code being
    # used for all PDU types/components.
    pdu_map = {
        "BASIC": "B-PDU",
        "METERED": "M-PDU",
        "SWITCHED": "S-PDU",
        "MANAGED": "MG-PDU",
    }

    # Each PDU model gets its own unique user-code suffix based on
    # its position within that PDU type in the Excel master.
    # Example: the 4 Basic PDU models become B-PDU1, B-PDU2,
    # B-PDU3 and B-PDU4. Metered/Switched/Managed follow the same rule.
    for part, qty in st.session_state.pdu_qty.items():
        if numeric(qty) <= 0:
            continue

        pdu_row = pdus_df[
            pdus_df["Part Code"].astype(str).str.strip() == str(part).strip()
        ]

        if pdu_row.empty:
            continue

        pdu_type = clean_text(pdu_row.iloc[0]["Type"]).upper()
        pdu_prefix = pdu_map.get(pdu_type)

        if not pdu_prefix:
            continue

        # Preserve Excel order so each model has a stable unique number.
        same_type = pdus_df[
            pdus_df["Type"].astype(str).str.strip().str.upper() == pdu_type
        ].reset_index(drop=True)

        matches = same_type.index[
            same_type["Part Code"].astype(str).str.strip() == str(part).strip()
        ].tolist()

        pdu_number = (matches[0] + 1) if matches else 1
        codes.append(f"{pdu_prefix}{pdu_number}")

    return "-".join(codes) if codes else "—"


def handle_excel_download():
    """
    Download callback.

    IMPORTANT:
    - User Count is NOT incremented here.
    - User Count belongs to the Streamlit session lifecycle.
    - Every Excel/PDF download creates one Configuration History entry.
    """
    # Generate the current user/configuration code for this download.
    st.session_state.user_code = generate_user_code()

    # Save this download to Configuration/User History.
    conn = sqlite3.connect(TRACKING_DB)

    conn.execute(
        """
        INSERT INTO user_history (
            user_code,
            customer_name,
            created_at
        )
        VALUES (?, ?, ?)
        """,
        (
            st.session_state.user_code,
            st.session_state.customer_name,
            datetime.now().isoformat(),
        ),
    )

    conn.commit()
    conn.close()


def render_user_info_panel(container):
    """Render the compact user information panel at its original location."""
    current_date = datetime.now().strftime("%d-%m-%Y")
    container.html(f"""
    <div style="
        background:#F7FBFF;
        border:1px solid #C9DFF2;
        border-radius:6px;
        padding:8px 14px;
        height:48px;
        display:flex;
        align-items:center;
    ">
        <div style="
            display:flex;
            width:100%;
            justify-content:space-between;
            align-items:center;
            text-align:center;
        ">
            <div style="flex:1;">
                <span style="font-size:9px;color:#64748B;font-weight:700;">USER CODE</span><br>
                <span style="font-size:13px;font-weight:700;color:#003B71;">
                    {st.session_state.user_code}
                </span>
            </div>
            <div style="flex:1;">
                <span style="font-size:9px;color:#64748B;font-weight:700;">USER COUNT</span><br>
                <span style="font-size:13px;font-weight:700;color:#003B71;">
                    {st.session_state.user_count}
                </span>
            </div>
            <div style="flex:1;">
                <span style="font-size:9px;color:#64748B;font-weight:700;">DATE</span><br>
                <span style="font-size:13px;font-weight:700;color:#003B71;">
                    {current_date}
                </span>
            </div>
        </div>
    </div>
    """)


# ============================================================
# DATA LOOKUPS
# ============================================================

def selected_config_record():
    match = configs_df[
        (configs_df["MDC Type"] == st.session_state.mdc_type)
        & (configs_df["Configuration"] == st.session_state.configuration)
    ]
    return match.iloc[0] if not match.empty else None


def selected_components():
    if st.session_state.mdc_type != "Single Rack":
        return pd.DataFrame(columns=components_df.columns)

    return components_df[
        (components_df["MDC Type"] == st.session_state.mdc_type)
        & (components_df["Configuration"] == st.session_state.configuration)
    ].copy()


def selected_multirack_components():
    """A is always included; B/C are included only when selected."""
    if st.session_state.mdc_type != "Multirack":
        return pd.DataFrame(columns=multirack_components_df.columns)

    rows = multirack_components_df[
        (multirack_components_df["MDC Type"] == "Multirack")
        & (multirack_components_df["Configuration"] == st.session_state.configuration)
    ].copy()
    if rows.empty:
        return rows

    selected_keys = {
        str(key) for key, selected in st.session_state.multirack_selected.items() if selected
    }
    is_a = rows["Group"].astype(str).str.upper().eq("A")
    is_selected_optional = rows["Selection Key"].astype(str).isin(selected_keys)
    return rows[is_a | is_selected_optional].copy()



# ============================================================
# BOM
# ============================================================

def build_bom():
    rows = []

    # --------------------------------------------------------
    # MULTIRACK / SINGLE-RACK BASE CONFIGURATION
    # --------------------------------------------------------
    # Multirack components come from the selected A/B/C group rows.
    # The current workbook has no Part Code/Cost, so those values remain
    # blank. If Part Code/Cost is added later, the parser reads them and
    # the same BOM logic uses them automatically.
    if st.session_state.mdc_type == "Multirack":
        for _, r in selected_multirack_components().iterrows():
            qty = numeric(r.get("Quantity"))
            unit_cost = numeric(r.get("Unit Cost"))
            part_code = clean_text(r.get("Part Code"))

            rows.append({
                "S.No.": clean_text(r.get("S.No.")),
                "Component Type": clean_text(r.get("Group")) or "Multirack",
                "Part Code": part_code,
                "Description": clean_text(r.get("Description")),
                "Quantity": float(qty) if pd.notna(qty) else "",
                "UOM": clean_text(r.get("UOM")) or "EA",
                "Unit Cost": unit_cost,
                "Total Cost": (
                    unit_cost * qty
                    if pd.notna(unit_cost) and pd.notna(qty)
                    else float("nan")
                ),
                "Unit Price": float("nan"),
                "Total Price": float("nan"),
                "Source": "Configuration",
            })

    else:
            # --------------------------------------------------------
            # SINGLE-RACK BASE CONFIGURATION
            # --------------------------------------------------------
        main_mdc_added = False

        for _, r in selected_components().iterrows():
            part_code = clean_text(r["Part Code"])

            if part_code == "801029209":
                if main_mdc_added:
                    continue
                main_mdc_added = True

            cost = numeric(r["Unit Cost"])
            qty = numeric(r["Quantity"])

            rows.append({
                "S.No.": len(rows) + 1,
                "Component Type": "Base (Configuration)",
                "Part Code": clean_text(r["Part Code"]),
                "Description": clean_text(r["Description"]),
                "Quantity": 0 if pd.isna(qty) else float(qty),
                "UOM": clean_text(r["UOM"]),
                "Unit Cost": cost,
                "Total Cost": (
                    cost * qty
                    if pd.notna(cost) and pd.notna(qty)
                    else float("nan")
                ),
                "Source": "Configuration",
            })

    # --------------------------------------------------------
    # PDU
    # --------------------------------------------------------
    for _, r in pdus_df.iterrows():
        part = clean_text(r["Part Code"])
        qty = numeric(st.session_state.pdu_qty.get(part, 0))

        if pd.notna(qty) and qty > 0:
            cost = numeric(r["Unit Cost"])

            c13_value = numeric(r["C13"])
            c19_value = numeric(r["C19"])
            c13_text = f"{c13_value:g}" if pd.notna(c13_value) else ""
            c19_text = f"{c19_value:g}" if pd.notna(c19_value) else ""

            desc = (
                f'{clean_text(r["Description"])} | '
                f'Type: {clean_text(r["Type"])} | '
                f'C13: {c13_text} | '
                f'C19: {c19_text}'
            )

            rows.append({
                "S.No.": len(rows) + 1,
                "Component Type": "PDU",
                "Part Code": part,
                "Description": desc,
                "Quantity": float(qty),
                "UOM": clean_text(r["UOM"]),
                "Unit Cost": cost,
                "Total Cost": cost * qty if pd.notna(cost) else float("nan"),
                "Source": "PDU",
            })

    # --------------------------------------------------------
    # OPTIONAL ACCESSORIES
    # --------------------------------------------------------
    accessory_lookup = {
        clean_text(r["Part Code"]): r
        for _, r in accessories_df.iterrows()
        if clean_text(r["Part Code"])
    }

    for part, selected_qty in st.session_state.accessory_qty.items():
        qty = numeric(selected_qty)

        if pd.isna(qty) or qty <= 0:
            continue

        r = accessory_lookup.get(str(part).strip())
        if r is None:
            continue

        cost = numeric(r["Unit Cost"])

        rows.append({
            "S.No.": len(rows) + 1,
            "Component Type": "Optional Accessory",
            "Part Code": clean_text(r["Part Code"]),
            "Description": clean_text(r["Description"]),
            "Quantity": float(qty),
            "UOM": clean_text(r["UOM"]),
            "Unit Cost": cost,
            "Total Cost": cost * qty if pd.notna(cost) else float("nan"),
            "Source": "Optional Accessory",
        })

    return pd.DataFrame(rows)


def cost_summary(bom):
    if bom.empty:
        return 0.0, 0.0, 0.0, 0.0

    base_cost = float(
        pd.to_numeric(
            bom.loc[bom["Source"] == "Configuration", "Total Cost"],
            errors="coerce",
        ).fillna(0).sum()
    )

    optional_cost = float(
        pd.to_numeric(
            bom.loc[bom["Source"] == "Optional Accessory", "Total Cost"],
            errors="coerce",
        ).fillna(0).sum()
    )

    pdu_cost = float(
        pd.to_numeric(
            bom.loc[bom["Source"] == "PDU", "Total Cost"],
            errors="coerce",
        ).fillna(0).sum()
    )

    return base_cost, optional_cost, pdu_cost, base_cost + optional_cost + pdu_cost


def add_selling_prices(bom, total_cost, margin_pct, freight, installation):
    """
    Keep the existing Single Rack pricing logic.

    For Multirack, A/B/C solution rows have blank Unit Cost when the
    Multirack workbook does not provide costing. PDU and Other Accessories
    still use the existing Single Rack pricing logic from MDC_Master_V1.xlsx.
    If Part Code/Cost is later added to the Multirack workbook, those
    solution rows are also handled automatically.
    """
    result = bom.copy()

    if result.empty:
        margin_price = 0.0
        final_selling_price = float(freight) + float(installation)
        return result, margin_price, final_selling_price

    result["Unit Price"] = pd.to_numeric(
        result["Unit Cost"], errors="coerce"
    )

    result["Total Price"] = (
        pd.to_numeric(result["Unit Price"], errors="coerce")
        * pd.to_numeric(result["Quantity"], errors="coerce")
    )

    margin_price = (
        total_cost / (1 - margin_pct / 100)
        if margin_pct < 100
        else 0.0
    )

    final_selling_price = (
        margin_price + float(freight) + float(installation)
    )

    return result, margin_price, final_selling_price


# ============================================================
# SALES FINAL BOQ VIEW
# ============================================================

def build_sales_boq(bom):
    """
    Compact sales/customer-facing BOQ.

    Display only:
        S.No. | Description | Quantity | UOM

    The original detailed BOM is NOT changed, so all individual
    component costs continue to be used for Final Selling Price.
    Cooling components are displayed as one line and camera
    components are displayed as one 4MP/1TB line.
    """
    if bom is None or bom.empty:
        return pd.DataFrame(
            columns=["S.No.", "Description", "Quantity", "UOM"]
        )

    # Multirack keeps the exact selected S.No./Description/Qty/UOM from
    # MULTIRACK BOQ.xlsx. The pricing columns remain available only in the
    # detailed Excel/PDF export and are intentionally blank.
    if st.session_state.mdc_type == "Multirack":
        output_rows = []

        for _, row in bom.iterrows():
            description = clean_text(row.get("Description"))
            if not description:
                continue

            qty = numeric(row.get("Quantity"))
            output_rows.append({
                "S.No.": clean_text(row.get("S.No.")),
                "Description": description,
                "Quantity": float(qty) if pd.notna(qty) else "",
                "UOM": clean_text(row.get("UOM")) or "EA",
            })

        return pd.DataFrame(
            output_rows,
            columns=["S.No.", "Description", "Quantity", "UOM"],
        )

    work = bom.copy()

    selected_config_components = selected_components()
    cooling_part_codes = set()

    if not selected_config_components.empty:
        cooling_rows = selected_config_components.tail(3)
        cooling_part_codes = set(
            cooling_rows["Part Code"]
            .dropna()
            .astype(str)
            .str.strip()
        )

    sales_rows = []
    cooling_source_rows = []
    camera_source_rows = []

    for idx, row in work.iterrows():
        component_type = clean_text(row.get("Component Type"))
        part_code = clean_text(row.get("Part Code"))
        description = clean_text(row.get("Description"))

        if (
            component_type == "Base (Configuration)"
            and part_code in cooling_part_codes
        ):
            cooling_source_rows.append(row)
            continue

        if (
            component_type == "Optional Accessory"
            and "CAMERA" in description.upper()
        ):
            camera_source_rows.append(row)
            continue

        sales_rows.append({
            "Description": description,
            "Quantity": numeric(row.get("Quantity")),
            "UOM": clean_text(row.get("UOM")) or "EA",
            "_source_index": idx,
        })

    # One compact Cooling Unit line.
    if cooling_source_rows:
        import re

        cooling_text = " ".join(
            clean_text(r.get("Description"))
            for r in cooling_source_rows
        )

        kw_matches = re.findall(
            r"(\d+(?:\.\d+)?)\s*(?:KW|K\.W\.)",
            cooling_text,
            flags=re.IGNORECASE,
        )

        cooling_type = (
            f"{kw_matches[0]} kW"
            if kw_matches
            else "Cooling Unit"
        )

        cooling_qty_values = [
            numeric(r.get("Quantity"))
            for r in cooling_source_rows
            if pd.notna(numeric(r.get("Quantity")))
            and numeric(r.get("Quantity")) > 0
        ]

        cooling_qty = max(cooling_qty_values) if cooling_qty_values else 1
        cooling_uom = clean_text(
            cooling_source_rows[0].get("UOM")
        ) or "EA"

        cooling_index = min(r.name for r in cooling_source_rows)

        insert_at = len(sales_rows)
        for pos, item in enumerate(sales_rows):
            if item["_source_index"] > cooling_index:
                insert_at = pos
                break

        sales_rows.insert(
            insert_at,
            {
                "Description": f"Cooling Unit {cooling_type}",
                "Quantity": cooling_qty,
                "UOM": cooling_uom,
                "_source_index": cooling_index,
            },
        )

    # One compact Camera line. Only 1TB camera rows are selectable in UI.
    if camera_source_rows:
        camera_qty_values = [
            numeric(r.get("Quantity"))
            for r in camera_source_rows
            if pd.notna(numeric(r.get("Quantity")))
            and numeric(r.get("Quantity")) > 0
        ]

        camera_qty = max(camera_qty_values) if camera_qty_values else 1
        camera_uom = clean_text(
            camera_source_rows[0].get("UOM")
        ) or "EA"
        camera_index = min(r.name for r in camera_source_rows)

        insert_at = len(sales_rows)
        for pos, item in enumerate(sales_rows):
            if item["_source_index"] > camera_index:
                insert_at = pos
                break

        sales_rows.insert(
            insert_at,
            {
                "Description": "Camera 4MP HDD 1TB",
                "Quantity": camera_qty,
                "UOM": camera_uom,
                "_source_index": camera_index,
            },
        )

    # --------------------------------------------------------
    # Selected configuration title from the Excel master.
    # This is the first row of the Sales Final BOQ.
    # Example:
    #   Configuration 1 -> SR1 3.5KW, W/O Dehumidifier
    #   Configuration 2 -> SR2 3.5KW, Dehumidifier
    # The title is already parsed from the corresponding
    # Solution block in MDC_Master_V1.xlsx by load_master().
    # --------------------------------------------------------
    selected_config = selected_config_record()
    selected_config_title = (
        clean_text(selected_config.get("Configuration Title", ""))
        if selected_config is not None
        else ""
    )

    # The selected configuration title is rendered as a centered,
    # full-width table row by the UI/Excel/PDF exporters. It is not
    # a numbered BOQ item. Component numbering therefore starts at 1.1.
    output_rows = []
    component_number = 1

    for item in sales_rows:
        description = clean_text(item["Description"])
        if not description:
            continue

        if selected_config_title:
            serial = f"1.{component_number}"
            component_number += 1
        else:
            # Fallback when the Excel configuration title is unavailable.
            serial = "1" if not output_rows else f"1.{component_number}"
            component_number += 1

        qty = numeric(item["Quantity"])

        output_rows.append({
            "S.No.": serial,
            "Description": description,
            "Quantity": float(qty) if pd.notna(qty) else "",
            "UOM": clean_text(item["UOM"]) or "EA",
        })

    return pd.DataFrame(
        output_rows,
        columns=["S.No.", "Description", "Quantity", "UOM", "_is_title"],
    ).drop(columns=["_is_title"], errors="ignore")


# ============================================================
# EXCEL / PDF OUTPUT
# ============================================================

def customer_table():
    # Export only the customer name.
    # Other customer details/problem/solution fields remain available in the UI
    # but are intentionally not included in Excel or PDF output.
    return pd.DataFrame([
        ["Customer Name", st.session_state.customer_name],
    ], columns=["Field", "Value"])


# ============================================================
# EXCEL EXPORT
# ============================================================

def excel_bytes(
    internal=False,
    bom=None,
    final_price=0.0,
    cost_data=None
):
    """
    Create ONE Excel worksheet containing:
    - Customer details
    - Final BOQ
    - Price summary
    - Final Selling Price

    Component Type is intentionally NOT shown.
    """

    output = BytesIO()

    if bom is None:
        bom = build_bom()

    # Single Rack Sales remains compact. Multirack exports use the same
    # detailed columns as the existing cost export, but pricing fields are
    # intentionally blank because the Multirack source has no pricing data.
    detailed_output = internal or st.session_state.mdc_type == "Multirack"
    display_bom = (
        bom
        if detailed_output
        else build_sales_boq(bom)
    )

    # --------------------------------------------------------
    # GET FINAL PRICE SAFELY
    # --------------------------------------------------------

    try:
        final_price = float(numeric(final_price))
    except Exception:
        final_price = 0.0

    # --------------------------------------------------------
    # EXCEL WRITER
    # --------------------------------------------------------

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        wb = writer.book

        # Create worksheet
        ws = wb.create_sheet("MDC BOQ")
        writer.sheets["MDC BOQ"] = ws

        # ====================================================
        # CUSTOMER / CONFIGURATION DETAILS
        # ====================================================

        info = customer_table()

        ws.cell(
            row=1,
            column=1,
            value="EATON MDC SOLUTION CONFIGURATOR"
        )

        ws.cell(
            row=2,
            column=1,
            value="Customer & Configuration Details"
        )

        row_no = 4

        for _, r in info.iterrows():

            ws.cell(
                row=row_no,
                column=1,
                value=r["Field"]
            )

            ws.cell(
                row=row_no,
                column=2,
                value=r["Value"]
            )

            row_no += 1

        # ====================================================
        # FINAL BOQ
        # ====================================================

        row_no += 1

        ws.cell(
            row=row_no,
            column=1,
            value="FINAL BOQ"
        )

        row_no += 1

        # ----------------------------------------------------
        # HEADERS
        # Component Type intentionally removed
        # ----------------------------------------------------

        if detailed_output:

            headers = [
                "S.No.",
                "Part Code",
                "Description",
                "Quantity",
                "UOM",
                "Unit Cost",
                "Total Cost",
                "Unit Price",
                "Total Price",
            ]

        else:

            headers = [
                "S.No.",
                "Description",
                "Quantity",
                "UOM",
            ]

        header_row = row_no

        for col_no, header in enumerate(
            headers,
            start=1
        ):

            ws.cell(
                row=header_row,
                column=col_no,
                value=header
            )

        row_no += 1

        # ----------------------------------------------------
        # SELECTED CONFIGURATION TITLE FROM EXCEL
        # ----------------------------------------------------
        selected_config = selected_config_record()
        selected_config_title = (
            clean_text(selected_config.get("Configuration Title", ""))
            if selected_config is not None
            else ""
        )

        if selected_config_title:
            ws.merge_cells(
                start_row=row_no,
                start_column=1,
                end_row=row_no,
                end_column=len(headers)
            )
            ws.cell(
                row=row_no,
                column=1,
                value=selected_config_title
            )
            ws.cell(
                row=row_no,
                column=1
            ).fill = PatternFill(
                "solid",
                fgColor="D9EAF7"
            )
            ws.cell(
                row=row_no,
                column=1
            ).font = Font(
                bold=True,
                color="003B71"
            )
            ws.cell(
                row=row_no,
                column=1
            ).alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            row_no += 1

        # ====================================================
        # BOQ ROWS
        # ====================================================

        for _, r in display_bom.iterrows():

            if detailed_output:

                values = [
                    r.get("S.No."),
                    r.get("Part Code"),
                    r.get("Description"),
                    r.get("Quantity"),
                    r.get("UOM"),
                    r.get("Unit Cost"),
                    r.get("Total Cost"),
                    r.get("Unit Price"),
                    r.get("Total Price"),
                ]

            else:

                values = [
                    r.get("S.No."),
                    r.get("Description"),
                    r.get("Quantity"),
                    r.get("UOM"),
                ]

            for col_no, value in enumerate(
                values,
                start=1
            ):

                if pd.isna(value):
                    value = None

                ws.cell(
                    row=row_no,
                    column=col_no,
                    value=value
                )

            row_no += 1

        # ====================================================
        # PRICE SUMMARY
        # ====================================================

        row_no += 1

        ws.cell(
            row=row_no,
            column=1,
            value="PRICE SUMMARY"
        )

        row_no += 1

        summary_start = row_no

        if detailed_output and st.session_state.mdc_type == "Multirack":

            summary = [
                ("Base Cost", None),
                ("Optional Cost", None),
                ("PDU Cost", None),
                ("Total Cost", None),
                ("Margin %", None),
                ("Margin Price", None),
                ("Freight", None),
                ("Installation", None),
                ("Warranty %", None),
                ("Warranty Amount", None),
                ("Final Selling Price", None),
            ]

        elif internal:

            summary = [
                ("Base Cost", base_cost),
                ("Optional Cost", optional_cost),
                ("PDU Cost", pdu_cost),
                ("Total Cost", total_cost),
                ("Margin %", margin_pct),
                ("Margin Price", margin_price),
                ("Freight", freight),
                ("Installation", installation),
                ("Warranty %", warranty_pct),
                ("Warranty Amount", margin_price * warranty_pct / 100),
                ("Final Selling Price", final_price),
            ]

        else:

            summary = [
                ("Final Selling Price", final_price),
            ]

        for label, value in summary:

            ws.cell(
                row=row_no,
                column=1,
                value=label
            )

            if value is None or (isinstance(value, float) and pd.isna(value)):
                numeric_value = None
            else:
                try:
                    numeric_value = float(numeric(value))
                except Exception:
                    numeric_value = None

            ws.cell(
                row=row_no,
                column=2,
                value=numeric_value
            )

            row_no += 1

        # ====================================================
        # EXCEL FORMATTING
        # ====================================================

        title_fill = "003B71"
        section_fill = "005EB8"
        header_fill = "D9EAF7"

        # ----------------------------------------------------
        # TITLE
        # ----------------------------------------------------

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=len(headers)
        )

        ws.merge_cells(
            start_row=2,
            start_column=1,
            end_row=2,
            end_column=len(headers)
        )

        for cell in ws[1]:

            cell.fill = PatternFill(
                "solid",
                fgColor=title_fill
            )

            cell.font = Font(
                color="FFFFFF",
                bold=True,
                size=16
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        for cell in ws[2]:

            cell.fill = PatternFill(
                "solid",
                fgColor=section_fill
            )

            cell.font = Font(
                color="FFFFFF",
                bold=True,
                size=11
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # ----------------------------------------------------
        # BOQ HEADER
        # ----------------------------------------------------

        for col_no in range(
            1,
            len(headers) + 1
        ):

            cell = ws.cell(
                row=header_row,
                column=col_no
            )

            cell.fill = PatternFill(
                "solid",
                fgColor=header_fill
            )

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        # ----------------------------------------------------
        # BOQ ALIGNMENT
        # ----------------------------------------------------

        boq_data_start = header_row + (2 if selected_config_title else 1)

        for rr in range(
            boq_data_start,
            boq_data_start + len(bom)
        ):

            # S.No.
            ws.cell(
                rr,
                1
            ).alignment = Alignment(
                horizontal="center"
            )

            if detailed_output:

                ws.cell(
                    rr,
                    4
                ).alignment = Alignment(
                    horizontal="center"
                )

                ws.cell(
                    rr,
                    5
                ).alignment = Alignment(
                    horizontal="center"
                )

                for cc in range(
                    6,
                    len(headers) + 1
                ):
                    ws.cell(
                        rr,
                        cc
                    ).alignment = Alignment(
                        horizontal="right"
                    )

            else:

                ws.cell(
                    rr,
                    1
                ).alignment = Alignment(
                    horizontal="center"
                )

                ws.cell(
                    rr,
                    3
                ).alignment = Alignment(
                    horizontal="center"
                )

                ws.cell(
                    rr,
                    4
                ).alignment = Alignment(
                    horizontal="center"
                )

        # ====================================================
        # CURRENCY FORMATTING
        # ====================================================

        # Excel BOQ currency columns
        if internal:

            currency_columns = [
                6,  # Unit Cost
                7,  # Total Cost
                8,  # Unit Price
                9,  # Total Price
            ]

        else:

            currency_columns = []

        boq_data_start = header_row + (2 if selected_config_title else 1)

        for rr in range(
            boq_data_start,
            boq_data_start + len(bom)
        ):

            for cc in currency_columns:

                ws.cell(
                    rr,
                    cc
                ).number_format = '₹ #,##0.00'

        # ====================================================
        # SUMMARY FORMATTING
        # ====================================================

        for rr in range(
            summary_start,
            row_no
        ):

            label = ws.cell(
                rr,
                1
            ).value

            value_cell = ws.cell(
                rr,
                2
            )

            value_cell.alignment = Alignment(
                horizontal="right"
            )

            if label in (
                "Margin %",
                "Warranty %"
            ):

                value_cell.number_format = '0.00'

            else:

                value_cell.number_format = (
                    '₹ #,##0.00'
                )

        # ----------------------------------------------------
        # FINAL SELLING PRICE HIGHLIGHT
        # ----------------------------------------------------

        for rr in range(
            summary_start,
            row_no
        ):

            if ws.cell(
                rr,
                1
            ).value == "Final Selling Price":

                ws.cell(
                    rr,
                    1
                ).font = Font(
                    bold=True,
                    color="003B71",
                    size=12
                )

                ws.cell(
                    rr,
                    2
                ).font = Font(
                    bold=True,
                    color="003B71",
                    size=12
                )

                ws.cell(
                    rr,
                    2
                ).number_format = (
                    '₹ #,##0.00'
                )

        # ====================================================
        # COLUMN WIDTHS
        # ====================================================

        if detailed_output:

            widths = {
                1: 10,
                2: 23,
                3: 65,
                4: 12,
                5: 10,
                6: 17,
                7: 17,
                8: 17,
                9: 17,
            }

        else:

            widths = {
                1: 10,
                2: 95,
                3: 12,
                4: 10,
            }

        for col_no, width in widths.items():

            ws.column_dimensions[
                get_column_letter(col_no)
            ].width = width

        # ====================================================
        # GENERAL SHEET SETTINGS
        # ====================================================

        ws.freeze_panes = (
            f"A{header_row + 1}"
        )

        last_boq_row = row_no - 1

        ws.auto_filter.ref = (
            f"A{header_row}:"
            f"{get_column_letter(len(headers))}"
            f"{last_boq_row}"
        )

        ws.sheet_view.showGridLines = False

        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 20
        ws.row_dimensions[header_row].height = 30

    output.seek(0)

    return output.getvalue()


# ============================================================
# PDF EXPORT
# ============================================================

# ============================================================
# PDF OUTPUT
# ============================================================

def pdf_bytes(internal=False, bom=None, final_price=0.0):
    """
    Create PDF report.

    Component Type is NOT displayed in the PDF.
    Final Selling Price is displayed clearly.
    """

    output = BytesIO()

    if bom is None:
        bom = build_bom()

    # Single Rack Sales remains compact. Multirack PDF uses the detailed
    # table so Part Code and all pricing columns are present but blank.
    detailed_output = internal or st.session_state.mdc_type == "Multirack"
    display_bom = bom if detailed_output else build_sales_boq(bom)

    # --------------------------------------------------------
    # FINAL PRICE
    # --------------------------------------------------------

    try:
        final_price = float(numeric(final_price))
    except Exception:
        final_price = 0.0

    # --------------------------------------------------------
    # PDF DOCUMENT
    # --------------------------------------------------------

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title="Eaton MDC Solution Configurator",
    )

    # --------------------------------------------------------
    # STYLES
    # --------------------------------------------------------

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "MdcTitle",
        parent=styles["Title"],
        fontSize=17,
        leading=20,
        alignment=TA_CENTER,
        spaceAfter=4,
    )

    sub_style = ParagraphStyle(
        "MdcSub",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        alignment=TA_CENTER,
        spaceAfter=8,
    )

    small = ParagraphStyle(
        "MdcSmall",
        parent=styles["Normal"],
        fontSize=7,
        leading=8,
    )

    center_small = ParagraphStyle(
        "MdcCenter",
        parent=small,
        alignment=TA_CENTER,
    )

    right_small = ParagraphStyle(
        "MdcRight",
        parent=small,
        alignment=TA_RIGHT,
    )

    # --------------------------------------------------------
    # STORY
    # --------------------------------------------------------

    story = [
        Paragraph(
            "EATON MDC SOLUTION CONFIGURATOR",
            title_style
        ),
        Paragraph(
            "Modular Data Center Solution Configuration & Pricing",
            sub_style
        ),
    ]

    # ========================================================
    # CUSTOMER / CONFIGURATION DETAILS
    # ========================================================

    info = customer_table()

    info_data = []

    for _, r in info.iterrows():

        field_value = clean_text(
            r.get("Field", "")
        )

        value_value = clean_text(
            r.get("Value", "")
        )

        info_data.append([
            Paragraph(
                f"<b>{field_value}</b>",
                small
            ),
            Paragraph(
                value_value,
                small
            ),
        ])

    info_table = Table(
        info_data,
        colWidths=[
            42 * mm,
            90 * mm
        ],
        hAlign="LEFT",
    )

    info_table.setStyle(
        TableStyle([
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.grey
            ),
            (
                "BACKGROUND",
                (0, 0),
                (0, -1),
                colors.HexColor("#D9EAF7")
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
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
                3
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                3
            ),
        ])
    )

    story.append(info_table)
    story.append(
        Spacer(1, 5 * mm)
    )

    # ========================================================
    # FINAL BOQ HEADING
    # ========================================================

    boq_heading_style = ParagraphStyle(
        "BOQHeading",
        parent=styles["Heading3"],
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#003B71"),
        spaceAfter=4,
        spaceBefore=0,
    )

    story.append(
        Paragraph(
            "FINAL BOQ",
            boq_heading_style
        )
    )

    # ========================================================
    # CUSTOMER PDF HEADERS
    #
    # IMPORTANT:
    # Component Type / Type is NOT included.
    # ========================================================

    if detailed_output:

        headers = [
            "S.No.",
            "Part Code",
            "Description",
            "Qty",
            "UOM",
            "Unit Cost",
            "Total Cost",
            "Unit Price",
            "Total Price",
        ]

    else:

        headers = [
            "S.No.",
            "Description",
            "Qty",
            "UOM",
        ]

    # ========================================================
    # BOQ TABLE DATA
    # ========================================================

    table_data = [headers]

    selected_config = selected_config_record()
    selected_config_title = (
        clean_text(selected_config.get("Configuration Title", ""))
        if selected_config is not None
        else ""
    )

    if selected_config_title:
        table_data.append([
            Paragraph(
                selected_config_title,
                ParagraphStyle(
                    "ConfigTitle",
                    parent=small,
                    fontSize=8,
                    leading=10,
                    alignment=TA_CENTER,
                    textColor=colors.HexColor("#003B71"),
                )
            )
        ] + [""] * (len(headers) - 1))

    for _, r in display_bom.iterrows():

        description = Paragraph(
            clean_text(
                r.get("Description", "")
            ),
            small
        )

        serial_no = Paragraph(
            clean_text(
                r.get("S.No.", "")
            ),
            center_small
        )

        quantity = Paragraph(
            clean_text(
                r.get("Quantity", "")
            ),
            center_small
        )

        uom = Paragraph(
            clean_text(
                r.get("UOM", "")
            ),
            center_small
        )

        if detailed_output:

            part_code = Paragraph(
                clean_text(
                    r.get("Part Code", "")
                ),
                small
            )

            row_data = [
                serial_no,
                part_code,
                description,
                quantity,
                uom,
                Paragraph(
                    money(r.get("Unit Cost")),
                    right_small
                ),
                Paragraph(
                    money(r.get("Total Cost")),
                    right_small
                ),
                Paragraph(
                    money(r.get("Unit Price")),
                    right_small
                ),
                Paragraph(
                    money(r.get("Total Price")),
                    right_small
                ),
            ]

        else:

            row_data = [
                serial_no,
                description,
                quantity,
                uom,
            ]

        table_data.append(row_data)

    # ========================================================
    # PDF COLUMN WIDTHS
    # ========================================================

    if detailed_output:

        widths = [
            12 * mm,   # S.No.
            30 * mm,   # Part Code
            78 * mm,   # Description
            13 * mm,   # Qty
            13 * mm,   # UOM
            25 * mm,   # Unit Cost
            27 * mm,   # Total Cost
            25 * mm,   # Unit Price
            27 * mm,   # Total Price
        ]

    else:

        widths = [
            13 * mm,   # S.No.
            135 * mm,  # Description
            18 * mm,   # Qty
            18 * mm,   # UOM
        ]

    # ========================================================
    # BOQ TABLE
    # ========================================================

    boq_table = Table(
        table_data,
        colWidths=widths,
        repeatRows=1,
        hAlign="CENTER",
    )

    # ========================================================
    # BOQ TABLE STYLE
    #
    # Explicit coordinates prevent column mismatch.
    # ========================================================

    boq_style_commands = [
        # Selected configuration title row
    ]

    if selected_config_title:
        boq_style_commands.extend([
            (
                "SPAN",
                (0, 1),
                (-1, 1)
            ),
            (
                "BACKGROUND",
                (0, 1),
                (-1, 1),
                colors.HexColor("#D9EAF7")
            ),
            (
                "TEXTCOLOR",
                (0, 1),
                (-1, 1),
                colors.HexColor("#003B71")
            ),
            (
                "FONTNAME",
                (0, 1),
                (-1, 1),
                "Helvetica-Bold"
            ),
            (
                "ALIGN",
                (0, 1),
                (-1, 1),
                "CENTER"
            ),
        ])

    boq_style_commands.extend([
        # Header
        (
            "BACKGROUND",
            (0, 0),
            (-1, 0),
            colors.HexColor("#003B71")
        ),

        (
            "TEXTCOLOR",
            (0, 0),
            (-1, 0),
            colors.white
        ),

        (
            "FONTNAME",
            (0, 0),
            (-1, 0),
            "Helvetica-Bold"
        ),

        (
            "FONTSIZE",
            (0, 0),
            (-1, -1),
            7
        ),

        # Grid
        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.3,
            colors.grey
        ),

        # Vertical alignment
        (
            "VALIGN",
            (0, 0),
            (-1, -1),
            "MIDDLE"
        ),

        # S.No.
        (
            "ALIGN",
            (0, 0),
            (0, -1),
            "CENTER"
        ),

        # Quantity/UOM/money alignment is added below based on output type.

        # Padding
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
    ])

    if detailed_output:
        boq_style_commands.extend([
            (
                "ALIGN",
                (3, 1),
                (4, -1),
                "CENTER"
            ),
            (
                "ALIGN",
                (5, 1),
                (-1, -1),
                "RIGHT"
            ),
        ])
    else:
        boq_style_commands.extend([
            (
                "ALIGN",
                (2, 1),
                (3, -1),
                "CENTER"
            ),
        ])

    boq_table.setStyle(
        TableStyle(
            boq_style_commands
        )
    )

    story.append(boq_table)

    story.append(
        Spacer(1, 5 * mm)
    )

    # ========================================================
    # PRICE SUMMARY
    # ========================================================

    if st.session_state.mdc_type == "Multirack":

        summary_data = [
            ["Base Cost", "", "Optional Cost", ""],
            ["PDU Cost", "", "Total Cost", ""],
            ["Margin %", "", "Margin Price", ""],
            ["Freight", "", "Installation", ""],
            ["Warranty %", "", "Warranty Amount", ""],
        ]

    elif internal:

        summary_data = [
            [
                "Base Cost",
                money(base_cost),
                "Optional Cost",
                money(optional_cost),
            ],
            [
                "PDU Cost",
                money(pdu_cost),
                "Total Cost",
                money(total_cost),
            ],
            [
                "Margin %",
                f"{margin_pct:.2f}%",
                "Margin Price",
                money(margin_price),
            ],
            [
                "Freight",
                money(freight),
                "Installation",
                money(installation),
            ],
            [
                "Warranty %",
                f"{warranty_pct:.2f}%",
                "Warranty Amount",
                money(
                    margin_price
                    * warranty_pct
                    / 100
                ),
            ],
        ]

        summary_table = Table(
            summary_data,
            colWidths=[
                38 * mm,
                42 * mm,
                45 * mm,
                45 * mm,
            ],
            hAlign="RIGHT",
        )

    else:

        summary_data = [
            [
                "FINAL SELLING PRICE",
                money(final_price)
            ]
        ]

        summary_table = Table(
            summary_data,
            colWidths=[
                55 * mm,
                45 * mm
            ],
            hAlign="RIGHT",
        )

    # ========================================================
    # SUMMARY STYLE
    # ========================================================

    summary_style = [
        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.4,
            colors.grey
        ),

        (
            "BACKGROUND",
            (0, 0),
            (-1, -1),
            colors.HexColor("#F4F8FC")
        ),

        (
            "FONTNAME",
            (0, 0),
            (-1, -1),
            "Helvetica-Bold"
        ),

        (
            "FONTSIZE",
            (0, 0),
            (-1, -1),
            8
        ),

        (
            "VALIGN",
            (0, 0),
            (-1, -1),
            "MIDDLE"
        ),

        (
            "ALIGN",
            (1, 0),
            (-1, -1),
            "RIGHT"
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
            5
        ),

        (
            "BOTTOMPADDING",
            (0, 0),
            (-1, -1),
            5
        ),
    ]

    summary_table.setStyle(
        TableStyle(summary_style)
    )

    story.append(summary_table)

    # ========================================================
    # FINAL SELLING PRICE
    #
    # Always show it clearly.
    # ========================================================

    story.append(
        Spacer(1, 4 * mm)
    )

    final_price_table = Table(
        [
            [
                "FINAL SELLING PRICE",
                "" if st.session_state.mdc_type == "Multirack" else money(final_price)
            ]
        ],
        colWidths=[
            55 * mm,
            45 * mm
        ],
        hAlign="RIGHT",
    )

    final_price_table.setStyle(
        TableStyle([
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.6,
                colors.HexColor("#003B71")
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, -1),
                colors.HexColor("#D9EAF7")
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, -1),
                colors.HexColor("#003B71")
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, -1),
                "Helvetica-Bold"
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                10
            ),

            (
                "ALIGN",
                (1, 0),
                (1, 0),
                "RIGHT"
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
                6
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                6
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                6
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                6
            ),
        ])
    )

    story.append(
        final_price_table
    )

    # ========================================================
    # BUILD PDF
    # ========================================================

    doc.build(story)

    output.seek(0)

    return output.getvalue()

# ============================================================

# CONSTANTS USED BY UI
# No product part numbers are hardcoded here.
# Product details are read from MDC_Master_V1.xlsx.


# HEADER
# ============================================================

# ============================================================
# LIGHT / COMPACT DESKTOP UI
# ============================================================
st.markdown("""
<style>
/* ---------- overall page ---------- */
[data-testid="stAppViewContainer"] {
    background: #F8FAFC;
}

[data-testid="stMainBlockContainer"] {
    max-width: 1500px;
    padding-top: 1.0rem;
    padding-bottom: 2rem;
}

/* ---------- title ---------- */
.mdc-title-card {
    background: linear-gradient(135deg, #075EA8 0%, #003B71 100%);
    border-radius: 12px;
    padding: 20px 26px 18px 26px;
    margin: 0 0 12px 0;
    box-shadow: 0 5px 18px rgba(0,59,113,.13);
}

.mdc-title {
    color: #FFFFFF !important;
    font-size: 30px !important;
    font-weight: 800 !important;
    letter-spacing: .15px;
    line-height: 1.15 !important;
    margin: 0 !important;
}

.mdc-subtitle {
    color: #DDEEFF !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    margin-top: 5px !important;
    line-height: 1.25 !important;
}

/* ---------- top information strip ---------- */
.mdc-meta-card {
    background: #FFFFFF;
    border: 1px solid #DCE7F2;
    border-radius: 10px;
    padding: 9px 14px;
    margin-bottom: 12px;
    box-shadow: 0 2px 8px rgba(15,23,42,.04);
}

.mdc-meta-grid {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    text-align: center;
}

.mdc-meta-item {
    flex: 1;
    min-width: 0;
}

.mdc-meta-label {
    color: #64748B;
    font-size: 9px;
    font-weight: 800;
    letter-spacing: .55px;
}

.mdc-meta-value {
    color: #003B71;
    font-size: 14px;
    font-weight: 800;
    margin-top: 2px;
}

/* ---------- section cards ---------- */
[data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid #DCE7F2 !important;
    border-radius: 11px !important;
    background: #FFFFFF !important;
    box-shadow: 0 2px 9px rgba(15,23,42,.035) !important;
}

.mdc-card-heading {
    display: flex;
    align-items: center;
    gap: 9px;
    color: #173B5E;
    font-size: 14px;
    font-weight: 800;
    letter-spacing: .15px;
    margin: 0 0 8px 0;
    padding-bottom: 7px;
    border-bottom: 1px solid #E8EFF6;
}

.mdc-number {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 23px;
    height: 23px;
    border-radius: 7px;
    background: #EAF3FB;
    color: #075EA8;
    font-size: 11px;
    font-weight: 800;
}

.mdc-mini-heading {
    color: #334E68;
    font-size: 11px;
    font-weight: 800;
    margin: 7px 0 3px 0;
    text-transform: uppercase;
    letter-spacing: .3px;
}

.mdc-help {
    color: #64748B;
    font-size: 10px;
    margin-top: -3px;
    margin-bottom: 5px;
}

/* ---------- compact Streamlit controls ---------- */
[data-testid="stWidgetLabel"] p {
    font-size: 11px !important;
    font-weight: 650 !important;
    color: #475569 !important;
    margin-bottom: 2px !important;
}

[data-testid="stTextInput"] input {
    font-size: 12px !important;
    min-height: 36px !important;
}

[data-testid="stNumberInput"] input {
    font-size: 12px !important;
}

[data-baseweb="select"] {
    font-size: 12px !important;
}

[data-baseweb="select"] * {
    font-size: 12px !important;
}

[data-testid="stRadio"] label,
[data-testid="stCheckbox"] label {
    font-size: 11px !important;
}

[data-testid="stRadio"] > div {
    gap: 8px !important;
}

[data-testid="stCheckbox"] {
    margin-bottom: -4px !important;
}

.stButton > button,
.stDownloadButton > button {
    min-height: 34px !important;
    font-size: 12px !important;
    border-radius: 7px !important;
}

/* ---------- small selected-detail panel ---------- */
.mdc-detail-box {
    background: #F7FAFD;
    border: 1px solid #E1EAF2;
    border-radius: 7px;
    padding: 7px 9px;
    margin-top: 5px;
}

.mdc-detail-title {
    color: #075EA8;
    font-size: 10px;
    font-weight: 800;
    margin-bottom: 2px;
}

.mdc-detail-text {
    color: #475569;
    font-size: 10px;
    line-height: 1.35;
}

/* ---------- accessory rows ---------- */
.mdc-accessory-row {
    padding: 2px 0;
}

/* Reduce default vertical gaps without changing functionality. */
[data-testid="stVerticalBlock"] {
    gap: .45rem;
}

@media (max-width: 900px) {
    .mdc-title { font-size: 25px !important; }
    .mdc-title-card { padding: 17px 20px 15px 20px; }
}
</style>
""", unsafe_allow_html=True)

st.html("""
<div class="mdc-title-card">
    <div class="mdc-title">Eaton MDC Solution Configurator</div>
    <div class="mdc-subtitle">Modular Data Center Solution Configuration &amp; Pricing</div>
</div>
""")

# ============================================================
# TOP INFORMATION BAR
# CUSTOMER NAME + USER INFORMATION
# ============================================================

top_customer_col, top_info_col = st.columns([2.2, 5.8], gap="small")

# ------------------------------------------------------------
# CUSTOMER NAME
# ------------------------------------------------------------
with top_customer_col:
    customer_name = st.text_input(
        "Customer Name",
        value=st.session_state.customer_name,
        key="customer_name_input",
        placeholder="Enter customer name",
        label_visibility="collapsed",
    )

st.session_state.customer_name = customer_name.strip()


# ------------------------------------------------------------
# USER CODE / COUNT / DATE
# ------------------------------------------------------------
with top_info_col:
    # The placeholder is rendered after all selection widgets are processed,
    # so USER CODE reflects the current UI selection immediately.
    user_info_placeholder = st.empty()


# ============================================================
# SIDEBAR ACCESS
# ============================================================
with st.sidebar:
    st.header("User Access")

    mode = st.radio(
        "Select User Type",
        ["Sales", "Internal – MDC"],
        index=0 if st.session_state.mode == "Sales" else 1,
    )

    if mode != st.session_state.mode:
        st.session_state.mode = mode

        if mode == "Sales":
            st.session_state.authenticated = False

        st.rerun()

    if mode == "Internal – MDC":
        if not st.session_state.authenticated:
            st.warning("Internal MDC access requires a password.")

            pwd = st.text_input(
                "MDC Password",
                type="password",
            )

            if st.button(
                "Unlock Internal Mode",
                use_container_width=True,
            ):
                if pwd == internal_password():
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
        else:
            st.success("Internal mode unlocked.")

            if st.button(
                "Lock Internal Mode",
                use_container_width=True,
            ):
                st.session_state.authenticated = False
                st.session_state.mode = "Sales"
                st.rerun()

is_internal = (
    st.session_state.mode == "Internal – MDC"
    and st.session_state.authenticated
)


# ============================================================
# MAIN CONFIGURATION AREA
# LEFT  = MDC + Multirack A/B/C + PDU
# RIGHT = Other Accessories
# ============================================================

main_left, main_right = st.columns(
    [1.0, 1.15],
    gap="medium"
)

# ============================================================
# LEFT SIDE
# ============================================================
with main_left:

    # ========================================================
    # 01. MDC TYPE & CONFIGURATION
    # ========================================================
    with st.container(border=True):

        st.markdown(
            '<div class="mdc-card-heading">'
            '<span class="mdc-number">01</span>'
            '<span>MDC TYPE & CONFIGURATION</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        mdc_type = st.radio(
            "MDC Type",
            ["Single Rack", "Multirack"],
            horizontal=True,
            index=(0 if st.session_state.mdc_type == "Single Rack" else 1),
        )

        if mdc_type != st.session_state.mdc_type:
            st.session_state.mdc_type = mdc_type
            st.session_state.configuration = "Configuration 1"
            st.session_state.accessory_qty = {}
            st.session_state.pdu_qty = {}
            st.session_state.multirack_selected = {}
            st.session_state.configuration_id = generate_configuration_id()
            st.session_state.configuration_saved = False
            st.rerun()

        available = configs_df[
            configs_df["MDC Type"] == st.session_state.mdc_type
        ].copy()

        labels = available["Configuration"].tolist()

        configuration_display_names = {
            "Configuration 1": "Configuration 1 - 1SR, 42U×800W×1200D, 3.5KW, W/O Dehumidifier",
            "Configuration 2": "Configuration 2 - 1SR, 42U×800W×1200D, 3.5KW, Dehumidifier",
            "Configuration 3": "Configuration 3 - 1SR, 42U×800W×1200D, 7KW, W/O Dehumidifier",
            "Configuration 4": "Configuration 4 - 1SR, 42U×800W×1200D, 7KW, Dehumidifier",
        }

        if labels:
            old_configuration = st.session_state.configuration

            selected_configuration = st.selectbox(
                "Select Configuration",
                labels,
                index=(
                    labels.index(old_configuration)
                    if old_configuration in labels else 0
                ),
                format_func=lambda x: (
                    configuration_display_names.get(x, x)
                    if st.session_state.mdc_type == "Single Rack"
                    else clean_text(
                        available.loc[
                            available["Configuration"] == x,
                            "Configuration Title"
                        ].iloc[0]
                    ) or x
                ),
                key="configuration_selection",
            )

            if selected_configuration != old_configuration:
                st.session_state.configuration = selected_configuration
                st.session_state.multirack_selected = {}
            else:
                st.session_state.configuration = selected_configuration

    # ========================================================
    # 02. MULTIRACK SOLUTION COMPONENTS
    # Read directly from MULTIRACK BOQ.xlsx.
    # A = main configuration; B/C = optional selections.
    # ========================================================
    if st.session_state.mdc_type == "Multirack":
        with st.container(border=True):
            st.markdown(
                '<div class="mdc-card-heading">'
                '<span class="mdc-number">02</span>'
                '<span>SOLUTION COMPONENTS</span>'
                '</div>',
                unsafe_allow_html=True,
            )

            all_mr_rows = multirack_components_df[
                (multirack_components_df["MDC Type"] == "Multirack")
                & (multirack_components_df["Configuration"] == st.session_state.configuration)
            ].copy()

            if all_mr_rows.empty:
                if not os.path.exists(MULTIRACK_FILE):
                    st.warning("MULTIRACK BOQ.xlsx was not found in the app folder.")
                else:
                    st.info(
                        f"No components were found for {st.session_state.configuration} "
                        f"in the Multirack Excel workbook."
                    )
            else:
                encountered_groups = []
                for group in all_mr_rows["Group"].astype(str):
                    if group not in encountered_groups:
                        encountered_groups.append(group)

                group_order = []
                for preferred in ["A", "B", "C"]:
                    if preferred in encountered_groups:
                        group_order.append(preferred)
                for group in encountered_groups:
                    if group not in group_order:
                        group_order.append(group)

                for group in group_order:
                    group_rows = all_mr_rows[
                        all_mr_rows["Group"].astype(str) == str(group)
                    ].copy()
                    if group_rows.empty:
                        continue

                    labels = (
                        group_rows["Group Label"].dropna().astype(str).str.strip().tolist()
                        if "Group Label" in group_rows.columns else []
                    )
                    group_label = next((x for x in labels if x), group)

                    st.markdown(
                        f'<div class="mdc-mini-heading" style="text-align:center;">'
                        f'{clean_text(group_label)}</div>',
                        unsafe_allow_html=True,
                    )

                    for _, mr in group_rows.iterrows():
                        selection_key = clean_text(mr.get("Selection Key"))
                        description = clean_text(mr.get("Description"))
                        qty = numeric(mr.get("Quantity"))
                        qty_text = f"{qty:g}" if pd.notna(qty) else ""
                        uom = clean_text(mr.get("UOM")) or "EA"
                        if not description:
                            continue

                        label = description
                        if qty_text:
                            label = f"{description} ({qty_text} {uom})"

                        if str(group).upper() == "A":
                            st.markdown(
                                f'<div style="padding:4px 6px;color:#334155;">'
                                f'<span style="font-weight:700;">✓</span>&nbsp;{label}'
                                f'</div>',
                                unsafe_allow_html=True,
                            )
                        else:
                            selected = st.checkbox(
                                label,
                                value=bool(st.session_state.multirack_selected.get(selection_key, False)),
                                key=f"mr_component_{selection_key}",
                            )
                            st.session_state.multirack_selected[selection_key] = selected
    # ========================================================
    # 03. PDU SELECTION
    # EXACTLY THE SAME PDU LOGIC AS SINGLE RACK
    # ========================================================
    with st.container(border=True):
        st.markdown(
            '<div class="mdc-card-heading">'
            '<span class="mdc-number">'
            + ("03" if st.session_state.mdc_type == "Multirack" else "02")
            + '</span>'
            '<span>PDU SELECTION</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        pdu_type_display = {
            "BASIC": "Basic PDU",
            "METERED": "Metered PDU",
            "SWITCHED": "Switched PDU",
            "MANAGED": "Managed PDU",
        }

        excel_pdu_types = []
        if not pdus_df.empty:
            excel_pdu_types = [
                x for x in (
                    pdus_df["Type"].astype(str).str.strip().str.upper().unique().tolist()
                ) if x
            ]

        pdu_types = ["None"] + [
            pdu_type_display.get(x, f"{x.title()} PDU")
            for x in excel_pdu_types
        ]

        pdu_col1, pdu_col2, pdu_col3 = st.columns(
            [1.0, 2.0, 0.55], gap="small"
        )

        with pdu_col1:
            previous_pdu_type = st.session_state.get(
                "pdu_type_selection", "None"
            )
            if previous_pdu_type not in pdu_types:
                previous_pdu_type = "None"

            selected_pdu_type = st.selectbox(
                "PDU Type",
                pdu_types,
                index=pdu_types.index(previous_pdu_type),
                key="pdu_type_selection",
            )

        selected_part = None

        with pdu_col2:
            if selected_pdu_type != "None":
                reverse_type = {
                    v: k for k, v in pdu_type_display.items()
                }
                excel_pdu_type = reverse_type.get(
                    selected_pdu_type,
                    selected_pdu_type.replace(" PDU", "").upper(),
                )

                filtered_pdus = pdus_df[
                    pdus_df["Type"].astype(str).str.strip().str.upper()
                    == excel_pdu_type
                ].copy()

                if not filtered_pdus.empty:
                    pdu_options = [
                        f'{clean_text(r["Part Code"])} — {clean_text(r["Description"])}'
                        for _, r in filtered_pdus.iterrows()
                    ]
                    old_selection = st.session_state.get("pdu_model_selection")
                    pdu_index = (
                        pdu_options.index(old_selection)
                        if old_selection in pdu_options else 0
                    )

                    selected_pdu = st.selectbox(
                        "Select PDU",
                        pdu_options,
                        index=pdu_index,
                        key="pdu_model_selection",
                    )
                    selected_row = filtered_pdus.iloc[pdu_options.index(selected_pdu)]
                    selected_part = clean_text(selected_row["Part Code"])
                else:
                    st.warning(
                        f"No {selected_pdu_type} options found in MDC_Master_V1.xlsx."
                    )

        with pdu_col3:
            if selected_part:
                current_pdu_qty = int(
                    numeric(st.session_state.pdu_qty.get(selected_part, 1))
                )
                pdu_quantity = st.number_input(
                    "Qty",
                    min_value=1,
                    max_value=999,
                    step=1,
                    value=current_pdu_qty,
                    key=f"pdu_quantity_{selected_part}",
                )
                st.session_state.pdu_qty = {selected_part: pdu_quantity}
            else:
                st.number_input(
                    "Qty",
                    min_value=1,
                    max_value=999,
                    value=1,
                    step=1,
                    disabled=True,
                    key="pdu_quantity_none",
                )
                st.session_state.pdu_qty = {}


# ============================================================
# RIGHT SIDE
# 04 OTHER ACCESSORIES FOR MULTIRACK / 03 FOR SINGLE RACK
# EXACTLY THE SAME EXCEL-DRIVEN ACCESSORY SELECTION
# ============================================================
with main_right:

    with st.container(border=True):
        st.markdown(
            '<div class="mdc-card-heading">'
            '<span class="mdc-number">'
            + ("04" if st.session_state.mdc_type == "Multirack" else "03")
            + '</span>'
            '<span>OTHER ACCESSORIES</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        optional_lookup = {
            clean_text(r["Part Code"]): r
            for _, r in accessories_df.iterrows()
            if clean_text(r["Part Code"])
        }

        def excel_optional_rows(keyword=None):
            if accessories_df.empty:
                return pd.DataFrame()
            if not keyword:
                return accessories_df.copy()
            key = str(keyword).upper()
            mask = (
                accessories_df["Part Code"].astype(str).str.upper().str.contains(key, na=False)
                | accessories_df["Description"].astype(str).str.upper().str.contains(key, na=False)
            )
            return accessories_df[mask].copy()

        def remove_rows(rows):
            for _, r in rows.iterrows():
                part = clean_text(r["Part Code"])
                if part:
                    st.session_state.accessory_qty.pop(part, None)

        def add_rows(rows, quantity=1):
            for _, r in rows.iterrows():
                part = clean_text(r["Part Code"])
                if part:
                    st.session_state.accessory_qty[part] = quantity

        # ----------------------------------------------------
        # FIRE SUPPRESSION
        # ----------------------------------------------------
        fire_rows = excel_optional_rows("FIRE")
        external_fire = fire_rows[
            fire_rows["Description"].astype(str).str.upper().str.contains("EXTERNAL", na=False)
        ].copy()
        internal_fire = fire_rows[
            ~fire_rows["Description"].astype(str).str.upper().str.contains("EXTERNAL", na=False)
        ].copy()

        fire_current = "None"
        if any(
            numeric(st.session_state.accessory_qty.get(p, 0)) > 0
            for p in external_fire["Part Code"].astype(str).str.strip()
        ):
            fire_current = "External"
        elif any(
            numeric(st.session_state.accessory_qty.get(p, 0)) > 0
            for p in internal_fire["Part Code"].astype(str).str.strip()
        ):
            fire_current = "Internal"

        fire_selection = st.radio(
            "Fire Suppression",
            ["None", "External", "Internal"],
            index=["None", "External", "Internal"].index(fire_current),
            horizontal=True,
            key="fire_suppression_selection",
        )

        remove_rows(external_fire)
        remove_rows(internal_fire)
        if fire_selection == "External":
            add_rows(external_fire, 1)
        elif fire_selection == "Internal":
            add_rows(internal_fire, 1)

        # ----------------------------------------------------
        # CAMERA
        # ----------------------------------------------------
        all_camera_rows = excel_optional_rows("CAMERA")
        four_tb_rows = (
            all_camera_rows[
                all_camera_rows["Description"].astype(str).str.upper().str.contains("4TB", na=False)
            ].copy()
            if not all_camera_rows.empty else pd.DataFrame()
        )
        remove_rows(four_tb_rows)

        camera_rows = (
            all_camera_rows[
                ~all_camera_rows["Description"].astype(str).str.upper().str.contains("4TB", na=False)
            ].copy()
            if not all_camera_rows.empty else pd.DataFrame()
        )
        camera_parts = (
            camera_rows["Part Code"].astype(str).str.strip().tolist()
            if not camera_rows.empty else []
        )
        camera_current = any(
            numeric(st.session_state.accessory_qty.get(p, 0)) > 0
            for p in camera_parts
        )

        cam_col1, cam_col2 = st.columns(
            [4.5, 1.15], gap="small", vertical_alignment="center"
        )
        with cam_col1:
            camera_selection = st.checkbox(
                "Camera",
                value=camera_current,
                key="camera_system_selection",
            )
        with cam_col2:
            if camera_selection:
                current_camera_qty = max(
                    [
                        int(numeric(st.session_state.accessory_qty.get(p, 0)))
                        for p in camera_parts
                        if numeric(st.session_state.accessory_qty.get(p, 0)) > 0
                    ] + [1]
                )
                camera_qty = st.number_input(
                    "Qty",
                    min_value=1,
                    max_value=999,
                    value=int(st.session_state.get("camera_quantity", current_camera_qty)),
                    step=1,
                    key="camera_quantity",
                )
            else:
                camera_qty = 0

        if camera_selection:
            add_rows(camera_rows, int(camera_qty))
        else:
            remove_rows(camera_rows)

        # ----------------------------------------------------
        # OTHER OPTIONAL ACCESSORIES
        # ----------------------------------------------------
        other_accessory_keywords = [
            ("KEYBOARD", "Rotating Keyboard Tray"),
            ("CABLE MANAGER", "Cable Manager"),
            ("TOP CABLE TRAY", "Top Cable Tray"),
            ("BRUSH PANEL", "Brush Panel"),
        ]

        for keyword, fallback_label in other_accessory_keywords:
            rows = excel_optional_rows(keyword)
            if rows.empty:
                continue

            for _, r in rows.iterrows():
                part = clean_text(r["Part Code"])
                description = clean_text(r["Description"])
                if not part:
                    continue

                current_qty = int(
                    numeric(st.session_state.accessory_qty.get(part, 0))
                )

                acc_col1, acc_col2 = st.columns(
                    [4.5, 1.15], gap="small", vertical_alignment="center"
                )
                with acc_col1:
                    selected = st.checkbox(
                        description if description else fallback_label,
                        value=current_qty > 0,
                        key=f"other_acc_{part}",
                    )

                with acc_col2:
                    if selected:
                        qty = st.number_input(
                            "Qty",
                            min_value=1,
                            max_value=999,
                            step=1,
                            value=current_qty if current_qty > 0 else 1,
                            key=f"other_qty_{part}",
                            label_visibility="collapsed",
                        )
                        st.session_state.accessory_qty[part] = qty
                    else:
                        st.session_state.accessory_qty.pop(part, None)


# ============================================================
# END MAIN CONFIGURATION AREA
# ============================================================


# LIVE USER CODE UPDATE
# ============================================================
# Generate the code after all current UI selections have been processed.
# The placeholder above keeps the panel in the same compact top position.
st.session_state.user_code = generate_user_code()
render_user_info_panel(user_info_placeholder)

# ============================================================
# 5. FINAL BOQ
# ============================================================

bom = build_bom()

base_cost, optional_cost, pdu_cost, total_cost = cost_summary(bom)

margin_pct = float(st.session_state.margin_pct)
freight = float(st.session_state.freight)
installation = float(st.session_state.installation)

if st.session_state.mdc_type == "Multirack":
    bom_with_price, margin_price, final_selling_price = add_selling_prices(
        bom,
        total_cost,
        margin_pct,
        freight,
        installation,
    )
elif not bom.empty:
    bom_with_price, margin_price, final_selling_price = add_selling_prices(
        bom,
        total_cost,
        margin_pct,
        freight,
        installation,
    )
else:
    bom_with_price = bom.copy()
    margin_price = (
        total_cost / (1 - margin_pct / 100)
        if margin_pct < 100
        else 0.0
    )
    final_selling_price = margin_price + freight + installation


st.html(f"""
<div style="
    display:flex;justify-content:space-between;align-items:center;gap:15px;
    background:#003B71;color:white;padding:7px 12px;border-radius:5px;
    margin:10px 0 8px 0;
">
    <div style="font-size:14px;font-weight:700;">5. FINAL BOQ</div>
    <div style="display:flex;align-items:center;gap:8px;white-space:nowrap;">
        <span style="font-size:11px;font-weight:700;">FINAL SELLING PRICE</span>
        <span style="font-size:16px;font-weight:700;">{money(final_selling_price)}</span>
    </div>
</div>
""")

# Compact Final BOQ is hidden until the user opens the dropdown.
with st.expander("▼ Show Final BOQ", expanded=False):

    if not bom.empty:

        sales_display_bom = build_sales_boq(bom_with_price)

        if not sales_display_bom.empty:

            html = """
            <style>
            .sales-boq-wrap {
                width:100%;
                overflow-x:auto;
            }

            .sales-boq-table {
                width:100%;
                border-collapse:collapse;
                table-layout:fixed;
                font-family:Arial,sans-serif;
                font-size:12px;
                border:1px solid #D9E1E8;
            }

            .sales-boq-table th {
                background:#F4F6F8;
                color:#555;
                font-weight:600;
                text-align:left;
                padding:7px 7px;
                border-bottom:1px solid #D9E1E8;
            }

            .sales-boq-table td {
                padding:6px 7px;
                border-bottom:1px solid #E5E7EB;
                color:#333;
                vertical-align:middle;
                overflow-wrap:anywhere;
            }

            .sales-boq-table .serial {
                width:9%;
                text-align:center !important;
            }

            .sales-boq-table .description {
                width:67%;
            }

            .sales-boq-table .quantity {
                width:12%;
                text-align:center !important;
            }

            .sales-boq-table .uom {
                width:12%;
                text-align:center !important;
            }

            .sales-boq-table .sales-boq-title-row td {
                background:#D9EAF7;
                color:#003B71 !important;
                font-weight:700;
                font-size:12px;
                text-align:center;
            }
            </style>

            <div class="sales-boq-wrap">
            <table class="sales-boq-table">
                <thead>
                    <tr>
                        <th class="serial">S.No.</th>
                        <th class="description">Description</th>
                        <th class="quantity">Qty</th>
                        <th class="uom">UOM</th>
                    </tr>
                </thead>
                <tbody>
            """

            # Configuration title: no serial number, centered across
            # the full table width. It is not counted as a component.
            selected_config = selected_config_record()
            selected_config_title = (
                clean_text(selected_config.get("Configuration Title", ""))
                if selected_config is not None
                else ""
            )
            if selected_config_title:
                html += f"""
                    <tr class="sales-boq-title-row">
                        <td colspan="4" style="
                            text-align:center;
                            font-weight:700;
                            color:#003B71;
                            background:#D9EAF7;
                        ">{selected_config_title}</td>
                    </tr>
                """

            for _, row in sales_display_bom.iterrows():

                serial_no = clean_text(row.get("S.No."))
                description = clean_text(row.get("Description"))
                quantity_value = numeric(row.get("Quantity"))
                quantity = (
                    f"{quantity_value:g}"
                    if pd.notna(quantity_value)
                    else ""
                )
                uom = clean_text(row.get("UOM"))

                html += f"""
                    <tr>
                        <td class="serial">{serial_no}</td>
                        <td class="description">{description}</td>
                        <td class="quantity">{quantity}</td>
                        <td class="uom">{uom}</td>
                    </tr>
                """

            html += """
                </tbody>
            </table>
            </div>
            """

            st.html(html)

        else:
            st.info("No sales BOQ items are available for the current selection.")

    else:
        st.info("No components selected for the current configuration.")


# ============================================================
# 6/7. INTERNAL COST & SELLING PRICE
# ============================================================

if is_internal:

    section_header("6. COST SUMMARY — INTERNAL ONLY")

    a, b, c, d = st.columns(4)

    if st.session_state.mdc_type == "Multirack":

        with a:
            price_box("Base Cost", float("nan"))

        with b:
            price_box("Optional Cost", float("nan"))

        with c:
            price_box("PDU Cost", float("nan"))

        with d:
            price_box("Total Cost", float("nan"))

    else:

        with a:
            price_box("Base Cost", base_cost)

        with b:
            price_box("Optional Cost", optional_cost)

        with c:
            price_box("PDU Cost", pdu_cost)

        with d:
            price_box("Total Cost", total_cost)

        section_header("7. COST TO SELLING PRICE — INTERNAL ONLY")

        p1, p2, p3, p4 = st.columns(4)

        with p1:
            st.session_state.margin_pct = st.number_input(
                "Margin (%)",
                min_value=0.0,
                max_value=99.0,
                value=float(st.session_state.margin_pct),
                step=0.5,
            )

        with p2:
            st.session_state.freight = st.number_input(
                "Freight",
                min_value=0.0,
                value=float(st.session_state.freight),
                step=500.0,
            )

        with p3:
            st.session_state.installation = st.number_input(
                "Installation",
                min_value=0.0,
                value=float(st.session_state.installation),
                step=500.0,
            )

        with p4:
            st.session_state.warranty_pct = st.number_input(
                "Warranty (%)",
                min_value=0.0,
                max_value=100.0,
                value=float(st.session_state.warranty_pct),
                step=0.5,
            )

        margin_pct = float(st.session_state.margin_pct)
        freight = float(st.session_state.freight)
        installation = float(st.session_state.installation)
        warranty_pct = float(st.session_state.warranty_pct)

        margin_price = (
            total_cost / (1 - margin_pct / 100)
            if margin_pct < 100
            else 0.0
        )

        final_selling_price = margin_price + freight + installation
        warranty_amount = margin_price * warranty_pct / 100

        a, b, c, d = st.columns(4)

        with a:
            price_box("Margin Price", margin_price)

        with b:
            price_box("After Freight", margin_price + freight)

        with c:
            price_box("Final Selling Price", final_selling_price)

        with d:
            price_box("Warranty Amount", warranty_amount)


# ============================================================
# 8. DOWNLOADS — EXCEL + PDF
# ============================================================

section_header("8. DOWNLOADS")

if not bom.empty:
    if st.session_state.mdc_type == "Multirack":
        internal_cost_data = [
            ["Base Cost", None],
            ["Optional Cost", None],
            ["PDU Cost", None],
            ["Total Cost", None],
            ["Margin %", None],
            ["Margin Price", None],
            ["Freight", None],
            ["Installation", None],
            ["Final Selling Price", None],
        ]
    else:
        internal_cost_data = [
            ["Base Cost", base_cost],
            ["Optional Cost", optional_cost],
            ["PDU Cost", pdu_cost],
            ["Total Cost", total_cost],
            ["Margin %", margin_pct],
            ["Margin Price", margin_price],
            ["Freight", freight],
            ["Installation", installation],
            ["Final Selling Price", final_selling_price],
        ]

    sales_excel = excel_bytes(
        internal=False,
        bom=bom_with_price,
        final_price=final_selling_price,
    )
    sales_pdf = pdf_bytes(
        internal=False,
        bom=bom_with_price,
        final_price=final_selling_price,
    )

    if is_internal:
        internal_excel = excel_bytes(
            internal=True,
            bom=bom_with_price,
            final_price=final_selling_price,
            cost_data=internal_cost_data,
        )
        internal_pdf = pdf_bytes(
            internal=True,
            bom=bom_with_price,
            final_price=final_selling_price,
        )

        st.markdown("**Internal – MDC**")
        i1, i2 = st.columns(2)

        with i1:
            st.download_button(
                "⬇️ Download Internal Excel",
                data=internal_excel,
                file_name="MDC_Internal_Cost.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                on_click=handle_excel_download,
            )

        with i2:
            st.download_button(
                "📄 Download Internal PDF",
                data=internal_pdf,
                file_name="MDC_Internal_Cost.pdf",
                mime="application/pdf",
                use_container_width=True,
                on_click=handle_excel_download,
            )

        st.markdown("**Sales**")

    s1, s2 = st.columns(2)

    with s1:
        st.download_button(
            "⬇️ Download Sales Excel",
            data=sales_excel,
            file_name="MDC_Sales_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            on_click=handle_excel_download,
        )

    with s2:
        st.download_button(
            "📄 Download Sales PDF",
            data=sales_pdf,
            file_name="MDC_Sales_Output.pdf",
            mime="application/pdf",
            use_container_width=True,
            on_click=handle_excel_download,
        )
else:
    st.info("Select a configuration with available BOM data before downloading.")


# ============================================================
# 9. USER HISTORY
# INTERNAL USERS ONLY
# ============================================================

if is_internal:

    section_header("9. CONFIGURATION HISTORY")

    conn = sqlite3.connect(TRACKING_DB)

    history_df = pd.read_sql_query(
        """
        SELECT
            user_code AS "User Code",
            customer_name AS "Customer Name"
        FROM user_history
        ORDER BY id DESC
        """,
        conn,
    )

    conn.close()

    if not history_df.empty:

        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True,
            height=220,
        )

    else:

        st.info("No user history available yet.")

# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption("Eaton MDC Solution Configurator")

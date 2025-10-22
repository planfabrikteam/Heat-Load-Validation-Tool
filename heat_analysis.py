# heat_analysis.py - VOLLSTÄNDIGE VERSION
import re
from pathlib import Path
import numpy as np
import pandas as pd
import pdfplumber
from xlsxwriter.utility import xl_rowcol_to_cell


# ------------------------------------------
# PARAMETER & KONFIGURATION
# ------------------------------------------

ALPHA = 0.05  # 95% central interval

# Bautile of interest
CODES_OF_INTEREST = {"AW", "IW", "IT", "AF", "DA", "DE", "FB"}

# Table columns of interest
HEADERS = [
    "OR", "Bauteil", "Anzahl", "Breite", "Laenge_Hoehe",
    "A_brutto", "A_abzug", "A_netto", "Grenzt_an",
    "Temp_diff", "U_wert", "Psi", "U_eq", "FTk_Watt",
]

# Early exclusions (names + underground floors)
EXCLUDED_ROOM_PATTERNS = [r"\bschacht\b", r"\btreppenhaus\b", r"\bkeller\b"]

# Category specification
ALL_CATEGORIES = [
    "Big Room",
    "Corner Room",
    "Exposed Room",
    "Internal Room",
    "Small Room w/t Outside Wall",
    "Wetcell",
    "Wetcell Mit Aussenbezug",
]

# Name edit
CATEGORY_CANON = {
    "Wetcell with Mit Aussenbezug": "Wetcell Mit Aussenbezug",
}

# DEFAULT ALL-PROJECTS BENCHMARKS
DEFAULT_BENCHMARKS = {
    "Big Room": {
        "lower_bound": 12.989, "upper_bound": 39.673, "mean": 24.208, "stddev": 6.992
    },
    "Corner Room": {
        "lower_bound": 16.89, "upper_bound": 55.036, "mean": 29.938, "stddev": 9.576
    },
    "Exposed Room": {
        "lower_bound": 3.857, "upper_bound": 59.377, "mean": 28.349, "stddev": 11.355
    },
    "Internal Room": {
        "lower_bound": 0.0, "upper_bound": 10.655, "mean": 2.8116, "stddev": 3.0327
    },
    "Small Room w/t Outside Wall": {
        "lower_bound": 6.437, "upper_bound": 38.003, "mean": 23.090, "stddev": 6.978
    },
    "Wetcell": {
        "lower_bound": 0.0, "upper_bound": 37.943, "mean": 8.942, "stddev": 11.021
    },
    "Wetcell Mit Aussenbezug": {
        "lower_bound": 13.023, "upper_bound": 68.297, "mean": 32.0555, "stddev": 16.528
    },
}


# ------------------------------------------
# HILFSFUNKTIONEN
# ------------------------------------------

def _skip_room_early(room_name: str | None, geschoss: int | None) -> bool:
    """Skip excluded rooms"""
    name_norm = (room_name or "").strip().casefold()
    if name_norm:
        for pat in EXCLUDED_ROOM_PATTERNS:
            if re.search(pat, name_norm):
                return True
    if geschoss is not None and geschoss < 0:
        return True
    return False

def to_float(s):
    """Parse localized numerics safely. Return None for blanks/dashes."""
    if s is None:
        return None
    s = str(s).strip()
    if s in {"", "-"}:
        return None
    s = s.replace("'", "").replace(" ", "")
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

HEADER_RE = re.compile(r"^Raum\s+\S+", re.MULTILINE)

def extract_room_code_and_name(text: str):
    """Parse: 'Raum <code> <name>'."""
    m = re.search(r"^Raum\s+(\S+)\s+(.+)$", text, re.MULTILINE)
    if not m:
        return None, None
    return m.group(1).strip(), m.group(2).strip()

def extract_room_area(text: str):
    m = re.search(r"Raumgrundfläche\s+A\s+([\d'.,]+)", text)
    return to_float(m.group(1)) if m else None

def extract_geschoss(text: str):
    m = re.search(r"Geschoss\s+(-?\d+)\b", text)
    return int(m.group(1)) if m else None

def extract_raumtemperatur(text: str):
    m = re.search(r"Raumtemperatur.*?([0-9]+(?:[.,][0-9]+)?)\s*°C", text, re.DOTALL)
    return to_float(m.group(1)) if m else None

# Normheizlast (W)
NORM_RE = re.compile(r"Normheizlast[^0-9]{0,20}([\d'.,]+)\s*[Ww]", re.IGNORECASE | re.DOTALL)
def extract_normheizlast(text: str):
    m = NORM_RE.search(text)
    return to_float(m.group(1)) if m else None

# Mindest-Luftwechselrate
MINDEST_LUFTWECHSEL_RE = re.compile(
    r"Mindest[- ]?Luftwechselrate.*?([0-9]+(?:[.,][0-9]+)?)\s*(?:1\s*/\s*h|h-?1|/h)?",
    re.IGNORECASE | re.DOTALL
)
def extract_mindest_luftwechselrate(text: str) -> float | None:
    m = MINDEST_LUFTWECHSEL_RE.search(text)
    return to_float(m.group(1)) if m else None

def iter_room_blocks(pages):
    """Yield page index lists for each room. A block may span multiple pages."""
    i, n = 0, len(pages)
    while i < n:
        txt = pages[i].extract_text() or ""
        if HEADER_RE.search(txt):
            blk = [i]
            i += 1
            while i < n and not HEADER_RE.search(pages[i].extract_text() or ""):
                blk.append(i)
                i += 1
            yield blk
        else:
            i += 1

# Table extraction
TABLE_CFG = dict(
    vertical_strategy="lines",
    horizontal_strategy="lines",
    intersection_y_tolerance=2,
    snap_x_tolerance=3,
)


# ------------------------------------------
# KLASSIFIZIERUNG
# ------------------------------------------

def _norm_name(s: str) -> str:
    return (s or "").strip().casefold()

def _normalize_orientation(val) -> str | None:
    if val is None:
        return None
    s = str(val).upper().replace(" ", "").replace("-", "").replace("/", "")
    filtered = "".join(ch for ch in s if ch in {"N", "O", "S", "W"})
    return filtered or None

# Bathroom detection
RE_WETCELL_ANY = re.compile(
    r"(?:\bwc\b|\btoilette\b|\bbad\b|\bbadezimmer\b|\bdusche\b|\bdu\b|\bbath\b|\bshower\b|\btoilet\b|\bbathroom\b)",
    re.IGNORECASE
)
def _is_wetcell_name(name_norm: str) -> bool:
    return bool(RE_WETCELL_ANY.search(name_norm or ""))

def summarize_rooms(df_elements: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse element rows to one row per (geschoss, room_code, room_name).
    Compute the metrics the classifier needs.
    """
    if df_elements.empty:
        return pd.DataFrame()

    df = df_elements.copy()
    df["bauteil"] = df["bauteil"].astype(str).str.strip().str.upper()
    df["orientation"] = df["orientation"].astype(str).str.strip()
    df["orient_norm"] = df["orientation"].apply(_normalize_orientation)
    df["iw_is_cold"] = df["bauteil"].eq("IW") & (pd.to_numeric(df["angrenzende_temp"], errors="coerce") < 20)

    grp_cols = ["geschoss", "room_code", "room_name"]
    rows = []

    for key, g in df.groupby(grp_cols, dropna=False):
        geschoss, room_code, room_name = key
        name_norm = _norm_name(room_name)

        g_aw = g[g["bauteil"].eq("AW")]
        aw_row_count = int(len(g_aw))
        aw_unique_orientations = int(g_aw["orient_norm"].dropna().nunique())
        iw_cold_count = int(g["iw_is_cold"].sum())

        def _first(series):
            s = series.dropna()
            return s.iloc[0] if s.size else None

        rows.append({
            "room_code": room_code,
            "room_name": room_name,
            "name_norm": name_norm,
            "geschoss": geschoss,
            "room_area": _first(g["room_area"]),
            "raumtemperatur": _first(g["raumtemperatur"]),
            "normheizlast": _first(g["normheizlast"]),
            "mindest_luftwechselrate": _first(g.get("mindest_luftwechselrate", pd.Series([None]))),
            "aw_row_count": aw_row_count,
            "aw_unique_orientations": aw_unique_orientations,
            "iw_cold_count": iw_cold_count,
        })

    return pd.DataFrame(rows)

# Classification rules
def classify_room(r: pd.Series) -> str:
    """
    1) Exclusions
    2) Wetcells (with/without Außenbezug via Mindest-Luftwechselrate)
    3) Big (>25 m²)
    4) Exposed (>=1 cold IW and room temp >= 21 °C)
    5) Small (area < 25) with one AW orientation
    6) Internal (no AW)
    7) Corner (>=2 AW and >1 orientation)
    """
    name_norm = r["name_norm"] or ""
    area = r["room_area"] or 0.0
    t_room = r["raumtemperatur"]
    rate = r.get("mindest_luftwechselrate")

    # 1) Exclusions
    if re.search(r"\bschacht\b", name_norm) or re.search(r"\btreppenhaus\b", name_norm) or re.search(r"\bkeller\b", name_norm):
        return "Excluded Room"

    # 2) Wetcells
    if _is_wetcell_name(name_norm):
        if rate is not None and rate > 0:
            return "Wetcell Mit Aussenbezug"
        else:
            return "Wetcell"

    # 3) Big
    if area > 25:
        return "Big Room"

    # 4) Exposed
    if r["iw_cold_count"] >= 1 and (t_room is not None) and (t_room >= 21):
        return "Exposed Room"

    # 5) Small with one AW orientation
    if area < 25 and r["aw_unique_orientations"] == 1:
        return "Small Room w/t Outside Wall"

    # 6) Internal
    if r["aw_row_count"] == 0:
        return "Internal Room"

    # 7) Corner
    if r["aw_row_count"] >= 2 and r["aw_unique_orientations"] > 1:
        return "Corner Room"

    return "Unclassified"


def categorize_and_validate_rooms(df_elements: pd.DataFrame) -> pd.DataFrame:
    rooms = summarize_rooms(df_elements)
    rooms = rooms[rooms["normheizlast"].notna()].copy()
    if rooms.empty:
        return rooms

    rooms["category"] = rooms.apply(classify_room, axis=1)
    rooms["category"] = rooms["category"].replace(CATEGORY_CANON)
    rooms["required_heat_per_m2"] = rooms["normheizlast"] / rooms["room_area"]

    # Special post-rule: drop 'Entree' if it classified as Corner Room
    mask_entree_corner = (rooms["name_norm"] == "entree") & (rooms["category"] == "Corner Room")
    rooms = rooms[~mask_entree_corner].copy()

    return rooms[[
        "room_code", "room_name", "geschoss",
        "room_area", "raumtemperatur",
        "normheizlast", "required_heat_per_m2",
        "mindest_luftwechselrate",
        "aw_row_count", "aw_unique_orientations", "iw_cold_count",
        "category"
    ]]


# ------------------------------------------
# PROJEKT STATISTIK
# ------------------------------------------

def build_project_stats(validated_df: pd.DataFrame) -> pd.DataFrame:
    """
    Per-category stats for THIS project.
    """
    rows = []
    for cat, s in validated_df.groupby("category")["required_heat_per_m2"]:
        vals = pd.to_numeric(s, errors="coerce").dropna()
        if vals.empty:
            continue
        lo_q = float(vals.quantile(ALPHA / 2.0, interpolation="linear"))
        hi_q = float(vals.quantile(1.0 - ALPHA / 2.0, interpolation="linear"))
        mean = float(vals.mean())
        std  = float(vals.std(ddof=1)) if len(vals) > 1 else 0.0
        rows.append({
            "category": cat,
            "lower_bound": lo_q,
            "upper_bound": hi_q,
            "mean": mean,
            "stddev": std,
        })
    return pd.DataFrame(rows).sort_values("category").reset_index(drop=True)


# ------------------------------------------
# EXCEL EXPORT (VOLLSTÄNDIG)
# ------------------------------------------

def export_to_excel(validated_df: pd.DataFrame, pdf_path: str) -> str:
    out_xlsx = f"{Path(pdf_path).stem}_heat_load_validation_results.xlsx"

    validated_df = validated_df.copy()
    validated_df["category"] = validated_df["category"].replace(CATEGORY_CANON)

    project_stats = build_project_stats(validated_df)

    rooms_columns = [
        "geschoss", "category", "room_name", "room_code",
        "room_area", "raumtemperatur", "normheizlast", "required_heat_per_m2",
        "Category's Lower Boundary", "Category's Upper Boundary",
        "Category's Mean", "Category's Standard Deviation",
        "Project's lower boundary per category", "Project's upper boundary per category",
        "Mean per category", "standard deviation", "validation_status",
    ]

    rooms_base = (
        validated_df
        .drop(columns=["aw_row_count", "aw_unique_orientations", "iw_cold_count"], errors="ignore")
        .copy()
    )
    for col in rooms_columns:
        if col not in rooms_base.columns:
            rooms_base[col] = ""
    rooms_base = rooms_base[rooms_columns]

    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_hdr   = wb.add_format({"bold": True})
        fmt_w0    = wb.add_format({"num_format": "0"})
        fmt_w_m2  = wb.add_format({"num_format": "0.0"})
        fmt_area  = wb.add_format({"num_format": "0.0"})
        fmt_green = wb.add_format({"font_color": "#006100", "bg_color": "#C6EFCE"})
        fmt_low   = wb.add_format({"font_color": "#005B96", "bg_color": "#B7DEE8"})
        fmt_high  = wb.add_format({"font_color": "#9C0006", "bg_color": "#F8CBAD"})

        # All-Projects Benchmarks Sheet
        apb_sheet = "All-Projects Benchmarks"
        legacy_rows = []
        for cat in ALL_CATEGORIES:
            d = DEFAULT_BENCHMARKS[cat]
            legacy_rows.append({
                "category": cat,
                "lower_bound": d["lower_bound"],
                "upper_bound": d["upper_bound"],
                "mean": d["mean"],
                "stddev": d["stddev"],
            })
        legacy_df = pd.DataFrame(legacy_rows, columns=["category", "lower_bound", "upper_bound", "mean", "stddev"])
        legacy_df.to_excel(writer, sheet_name=apb_sheet, index=False)
        ws_apb = writer.sheets[apb_sheet]
        ws_apb.set_row(0, None, fmt_hdr)
        ws_apb.add_table(0, 0, len(legacy_df), 4, {
            "name": "AllProjectsBenchmarks",
            "columns": [{"header": h} for h in legacy_df.columns]
        })
        reserve = max(len(ALL_CATEGORIES), 200)
        wb.define_name("Legacy_Cat",  f"='{apb_sheet}'!$A$2:$A${1+reserve}")
        wb.define_name("Legacy_Low",  f"='{apb_sheet}'!$B$2:$B${1+reserve}")
        wb.define_name("Legacy_Up",   f"='{apb_sheet}'!$C$2:$C${1+reserve}")
        wb.define_name("Legacy_Mean", f"='{apb_sheet}'!$D$2:$D${1+reserve}")
        wb.define_name("Legacy_Std",  f"='{apb_sheet}'!$E$2:$E${1+reserve}")

        # ProjectStats Sheet
        project_stats.to_excel(writer, sheet_name="ProjectStats", index=False)
        ws_proj = writer.sheets["ProjectStats"]
        ws_proj.set_row(0, None, fmt_hdr)
        ws_proj.add_table(0, 0, max(len(project_stats), 1), 4, {
            "name": "ProjectStats",
            "columns": [{"header": h} for h in project_stats.columns]
        })
        p_rows = max(len(project_stats), 1)
        p_res  = max(p_rows, 200)
        wb.define_name("Proj_Cat",  f"=ProjectStats!$A$2:$A${1+p_res}")
        wb.define_name("Proj_Low",  f"=ProjectStats!$B$2:$B${1+p_res}")
        wb.define_name("Proj_Up",   f"=ProjectStats!$C$2:$C${1+p_res}")
        wb.define_name("Proj_Mean", f"=ProjectStats!$D$2:$D${1+p_res}")
        wb.define_name("Proj_Std",  f"=ProjectStats!$E$2:$E${1+p_res}")

        # Custom Intervals Sheet
        ci_sheet = "Custom Intervals"
        mo_df = pd.DataFrame({
            "category": ALL_CATEGORIES,
            "Manual lower boundary": ["" for _ in ALL_CATEGORIES],
            "Manual upper boundary": ["" for _ in ALL_CATEGORIES],
        })
        mo_df.to_excel(writer, sheet_name=ci_sheet, index=False)
        ws_ci = writer.sheets[ci_sheet]
        ws_ci.set_row(0, None, fmt_hdr)
        ws_ci.add_table(0, 0, len(mo_df), 2, {
            "name": "CustomIntervals",
            "columns": [{"header": h} for h in mo_df.columns]
        })
        m_res = max(len(ALL_CATEGORIES), 200)
        wb.define_name("Man_Cat", f"='{ci_sheet}'!$A$2:$A${1+m_res}")
        wb.define_name("Man_Low", f"='{ci_sheet}'!$B$2:$B${1+m_res}")
        wb.define_name("Man_Up",  f"='{ci_sheet}'!$C$2:$C${1+m_res}")

        # Rooms Sheet
        rooms_base.to_excel(writer, sheet_name="Rooms", index=False)
        ws = writer.sheets["Rooms"]
        ws.set_row(0, None, fmt_hdr)
        ws.freeze_panes(1, 0)

        n_rows = len(rooms_base)
        n_cols = len(rooms_base.columns)

        ws.add_table(0, 0, n_rows, n_cols - 1, {
            "name": "Rooms",
            "columns": [{"header": h} for h in rooms_base.columns]
        })

        # Column formats
        col_idx = {c: i for i, c in enumerate(rooms_base.columns)}
        def _setc(name, width, fmt):
            if name in col_idx: ws.set_column(col_idx[name], col_idx[name], width, fmt)
        _setc("room_area", 12, fmt_area)
        _setc("normheizlast", 12, fmt_w0)
        _setc("required_heat_per_m2", 16, fmt_w_m2)
        _setc("Category's Mean", 16, fmt_w_m2)
        _setc("Category's Standard Deviation", 16, fmt_w_m2)
        _setc("Project's upper boundary per category", 16, fmt_w_m2)
        _setc("Project's lower boundary per category", 16, fmt_w_m2)
        _setc("Mean per category", 16, fmt_w_m2)
        _setc("standard deviation", 16, fmt_w_m2)

        # Positions for formulas
        c_cat     = col_idx["category"]
        c_req     = col_idx["required_heat_per_m2"]
        c_catL    = col_idx["Category's Lower Boundary"]
        c_catU    = col_idx["Category's Upper Boundary"]
        c_catM    = col_idx["Category's Mean"]
        c_catS    = col_idx["Category's Standard Deviation"]
        c_prjL    = col_idx["Project's lower boundary per category"]
        c_prjU    = col_idx["Project's upper boundary per category"]
        c_prjM    = col_idx["Mean per category"]
        c_prjS    = col_idx["standard deviation"]
        c_status  = col_idx["validation_status"]

        for r in range(1, n_rows + 1):
            cat_cell  = xl_rowcol_to_cell(r, c_cat)
            req_cell  = xl_rowcol_to_cell(r, c_req)

            ws.write_formula(r, c_catL, f'=IFERROR(INDEX(Legacy_Low,MATCH({cat_cell},Legacy_Cat,0)),"")')
            ws.write_formula(r, c_catU, f'=IFERROR(INDEX(Legacy_Up,MATCH({cat_cell},Legacy_Cat,0)),"")')
            ws.write_formula(r, c_catM, f'=IFERROR(INDEX(Legacy_Mean,MATCH({cat_cell},Legacy_Cat,0)),"")')
            ws.write_formula(r, c_catS, f'=IFERROR(INDEX(Legacy_Std,MATCH({cat_cell},Legacy_Cat,0)),"")')

            ws.write_formula(r, c_prjL, f'=IFERROR(INDEX(Proj_Low,MATCH({cat_cell},Proj_Cat,0)),"")')
            ws.write_formula(r, c_prjU, f'=IFERROR(INDEX(Proj_Up,MATCH({cat_cell},Proj_Cat,0)),"")')
            ws.write_formula(r, c_prjM, f'=IFERROR(INDEX(Proj_Mean,MATCH({cat_cell},Proj_Cat,0)),"")')
            ws.write_formula(r, c_prjS, f'=IFERROR(INDEX(Proj_Std,MATCH({cat_cell},Proj_Cat,0)),"")')

            man_low = f'INDEX(Man_Low,MATCH({cat_cell},Man_Cat,0))'
            man_up  = f'INDEX(Man_Up,MATCH({cat_cell},Man_Cat,0))'
            leg_low = f'INDEX(Legacy_Low,MATCH({cat_cell},Legacy_Cat,0))'
            leg_up  = f'INDEX(Legacy_Up,MATCH({cat_cell},Legacy_Cat,0))'
            prj_low = f'INDEX(Proj_Low,MATCH({cat_cell},Proj_Cat,0))'
            prj_up  = f'INDEX(Proj_Up,MATCH({cat_cell},Proj_Cat,0))'

            ws.write_formula(
                r, c_status,
                'IF('
                f'AND(ISNUMBER({man_low}),ISNUMBER({man_up})),'
                f'IF({req_cell}<MIN({man_low},{man_up}),"Low",'
                f'IF({req_cell}>MAX({man_low},{man_up}),"High","Accepted")),'
                f'IF(AND(ISNUMBER({leg_low}),ISNUMBER({leg_up})),'
                f'IF({req_cell}<MIN({leg_low},{leg_up}),"Low",'
                f'IF({req_cell}>MAX({leg_low},{leg_up}),"High","Accepted")),'
                f'IF(AND(ISNUMBER({prj_low}),ISNUMBER({prj_up})),'
                f'IF({req_cell}<MIN({prj_low},{prj_up}),"Low",'
                f'IF({req_cell}>MAX({prj_low},{prj_up}),"High","Accepted")),'
                '"No baseline")))'
            )

        if n_rows > 0:
            ws.conditional_format(1, c_status, n_rows, c_status,
                                  {"type": "text", "criteria": "containing", "value": "Accepted", "format": fmt_green})
            ws.conditional_format(1, c_status, n_rows, c_status,
                                  {"type": "text", "criteria": "containing", "value": "Low", "format": fmt_low})
            ws.conditional_format(1, c_status, n_rows, c_status,
                                  {"type": "text", "criteria": "containing", "value": "High", "format": fmt_high})

        # Summary Sheet
        cats_present = sorted(validated_df["category"].dropna().unique())
        summary_categories = pd.DataFrame({"category": cats_present})
        summary_categories.to_excel(writer, sheet_name="Summary", index=False)
        ws_sum = writer.sheets["Summary"]
        ws_sum.set_row(0, None, fmt_hdr)
        ws_sum.write(0, 1, "rooms_total", fmt_hdr)
        ws_sum.write(0, 2, "rooms_accepted", fmt_hdr)
        ws_sum.write(0, 3, "rooms_rejected", fmt_hdr)

        for i in range(len(cats_present)):
            rr = i + 1
            ws_sum.write_formula(rr, 1, f'=COUNTIF(Rooms[category], A{rr+1})')
            ws_sum.write_formula(rr, 2, f'=COUNTIFS(Rooms[category], A{rr+1}, Rooms[validation_status], "Accepted")')
            ws_sum.write_formula(rr, 3, f'=B{rr+1}-C{rr+1}')

        ws_sum.add_table(0, 0, len(cats_present), 3, {
            "name": "Summary",
            "columns": [{"header": "category"},
                        {"header": "rooms_total"},
                        {"header": "rooms_accepted"},
                        {"header": "rooms_rejected"}]
        })

    return out_xlsx


# ------------------------------------------
# HAUPTPIPELINE
# ------------------------------------------

def run_pipeline(pdf_path: str):
    """
    End-to-end: extract -> classify -> export Excel with formulas and baselines.
    Returns: (validated_df, out_xlsx_path)
    """
    print("🔍 Extracting data from PDF...")
    rows = []

    with pdfplumber.open(pdf_path) as pdf:
        for blk in iter_room_blocks(pdf.pages):
            texts = [pdf.pages[p].extract_text() or "" for p in blk]
            full_txt = "\n".join(texts)

            # Room-level
            room_code, room_name = extract_room_code_and_name(full_txt)
            room_area = extract_room_area(full_txt)
            geschoss = extract_geschoss(full_txt)
            raumtemperatur = extract_raumtemperatur(full_txt)
            normheizlast = extract_normheizlast(full_txt)
            mindest_luftwechselrate = extract_mindest_luftwechselrate(full_txt)

            if _skip_room_early(room_name, geschoss):
                continue

            # Collect tables
            tables = []
            for p in blk:
                t = pdf.pages[p].extract_table(table_settings=TABLE_CFG)
                if t:
                    tables.append(t)

            if not tables or room_name is None or room_area is None:
                continue

            # Flatten and keep relevant element codes
            dfs = []
            for t in tables:
                if len(t) < 3:  # data starts after 2 header rows
                    continue
                df = pd.DataFrame(t[2:], columns=HEADERS)
                df = df[df["Bauteil"].notna()]
                df = df[df["Bauteil"].str.upper() != "BAUTEIL"]
                df = df[df["Bauteil"].str.strip().str.upper().isin(CODES_OF_INTEREST)]
                if not df.empty:
                    dfs.append(df)
            if not dfs:
                continue

            df_room = pd.concat(dfs, ignore_index=True)

            for _, r in df_room.iterrows():
                rows.append({
                    "room_code": room_code,
                    "room_name": room_name,
                    "room_area": room_area,
                    "geschoss": geschoss,
                    "raumtemperatur": raumtemperatur,
                    "orientation": r["OR"],
                    "bauteil": r["Bauteil"],
                    "angrenzende_temp": to_float(r["Temp_diff"]),
                    "normheizlast": normheizlast,
                    "mindest_luftwechselrate": mindest_luftwechselrate,
                })

    df_all = pd.DataFrame(rows)
    uniq_rooms = df_all[["geschoss", "room_code", "room_name"]].drop_duplicates().shape[0]
    print(f"📊 {len(df_all)} elements from {uniq_rooms} unique rooms")

    results = categorize_and_validate_rooms(df_all)
    if results.empty:
        print("❌ Error: No rooms with valid normheizlast values found!")
        return None, None

    results_sorted = (
        results.copy()
        .sort_values(by=["geschoss", "room_name", "room_code"],
                     ascending=[True, True, True],
                     kind="mergesort")
        .reset_index(drop=True)
    )

    out_xlsx = export_to_excel(results_sorted, pdf_path)
    print(f"\n✅ Results saved to: {out_xlsx}")
    return results_sorted, out_xlsx

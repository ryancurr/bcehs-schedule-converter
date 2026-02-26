import datetime as dt
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook


# ===================== SHARED RULES (ACP + PCP) =====================

STUDENT_ALIASES = {
    "rory": "Rory-lynn Bradshaw",
}

EXCLUDE_STUDENTS = {
    "jadyn",
    "jadyn langley",
}

PARTNER_MARKERS = {
    "partner", "parnter", "partnre", "parter", "prtnr",
}

COLUMBIA_SUFFIX_PAT = re.compile(r"\s*-\s*Columbia\s*$", re.I)

# ===================================================================

TIME_RANGE_PAT = re.compile(
    r"(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})|(\d{3,4})\s*[-–]\s*(\d{3,4})"
)
CODE_TOKEN_PAT = re.compile(r"^\d{3}[A-Za-z0-9]+$")
AMBULANCE_PAT = re.compile(r"^(\d{3}[A-Z]\d)", re.I)
HEADER_DATE_PAT = re.compile(r"\b([A-Za-z]{3})/(\d{1,2})\b")

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def bytes_to_filelike(b: bytes):
    import io
    return io.BytesIO(b)


def norm_hhmm(x: str) -> str:
    x = str(x)
    if len(x) == 3:
        h = int(x[0])
        m = int(x[1:])
    else:
        h = int(x[:-2])
        m = int(x[-2:])
    return f"{h:02d}:{m:02d}"


def parse_shift(text: str) -> Optional[dict]:
    if not isinstance(text, str):
        return None

    t = text.replace("\n", " ").strip()
    if not t or t == "\\":
        return None

    code = ""
    for tok in t.split():
        if CODE_TOKEN_PAT.match(tok):
            code = tok
            break

    start = end = ""
    m = TIME_RANGE_PAT.search(t)
    if m:
        if m.group(1):
            start, end = m.group(1), m.group(2)
        else:
            start, end = norm_hhmm(m.group(3)), norm_hhmm(m.group(4))

    if not (code or start or end):
        return None

    station = code[:3] if code else ""
    ambulance = ""
    if code:
        m2 = AMBULANCE_PAT.match(code)
        if m2:
            ambulance = m2.group(1)

    return {
        "raw": t,
        "code": code,
        "start": start,
        "end": end,
        "station": station,
        "ambulance": ambulance,
    }


def format_preceptor_one(name: str) -> str:
    s = re.sub(r"\s+", " ", str(name).strip())
    if "," in s:
        last, rest = s.split(",", 1)
        return re.sub(r"\s+", " ", f"{rest.strip()} {last.strip()}".strip())
    return s


def format_preceptor(name: str) -> str:
    if not isinstance(name, str):
        return ""
    parts = [p.strip() for p in name.split("/") if p.strip()]
    parts = [format_preceptor_one(p) for p in parts]
    return " / ".join(parts)


def is_partner_marker(s: str) -> bool:
    if not isinstance(s, str):
        return False
    t = s.strip().lower()
    if t in PARTNER_MARKERS:
        return True
    return ("partner" in t) or ("parnter" in t)


def normalize_student(raw: str) -> str:
    if not isinstance(raw, str):
        return ""

    s = re.sub(r"\s+", " ", raw.strip())
    if not s:
        return ""

    low = s.lower()
    if low in {"student", "n/a", "na"}:
        return ""
    if is_partner_marker(s):
        return ""

    # Strip "- Columbia" if present (but DO NOT require it anymore)
    s = re.sub(COLUMBIA_SUFFIX_PAT, "", s).strip()

    if s.lower() in STUDENT_ALIASES:
        s = STUDENT_ALIASES[s.lower()]

    if s.lower() in EXCLUDE_STUDENTS:
        return ""

    return s


def parse_header_dates(ws, year: int, start_col: int) -> Dict[int, dt.date]:
    col_dates: Dict[int, dt.date] = {}
    for c in range(start_col, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str):
            m = HEADER_DATE_PAT.search(v)
            if m:
                mon_abbr = m.group(1).lower()
                day = int(m.group(2))
                month = MONTHS.get(mon_abbr)
                if month:
                    col_dates[c] = dt.date(year, month, day)
    return col_dates


# ===================== PCP =====================

def extract_pcp_rows(wb, year: int) -> pd.DataFrame:
    rows: List[dict] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        col_dates = parse_header_dates(ws, year, start_col=2)
        if not col_dates:
            continue

        for r in range(2, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if not isinstance(a, str):
                continue

            a_str = a.strip()
            if not a_str or a_str.upper().startswith("STUDENT") or a_str == "Preceptor":
                continue

            preceptor = format_preceptor_one(a_str)

            up = ws.cell(r - 1, 1).value
            if not (isinstance(up, str) and up.strip().upper().startswith("STUDENT")):
                continue

            for c, date in col_dates.items():
                v = ws.cell(r, c).value
                if not isinstance(v, str):
                    continue

                sh = parse_shift(v)
                if not sh:
                    continue

                student_raw = ws.cell(r - 1, c).value
                student = normalize_student(student_raw)
                if not student:
                    continue

                rows.append({
                    "Student Name": student,
                    "Date (YYYY-MM-DD)": date.isoformat(),
                    "Start Time (HH:MM)": sh["start"],
                    "End Time (HH:MM)": sh["end"],
                    "Location": sheet_name,
                    "Station": sh["station"],
                    "Ambulance Number": sh["ambulance"],
                    "Preceptor": preceptor,
                })

    return pd.DataFrame(rows)


# ===================== ACP =====================

def is_student_marker(v) -> bool:
    return isinstance(v, str) and v.strip().upper().startswith("STUDENT")


def extract_acp_rows(wb, year: int) -> pd.DataFrame:
    ws = wb[wb.sheetnames[0]]
    col_dates = parse_header_dates(ws, year, start_col=3)

    rows: List[dict] = []
    pending_student_rows: List[int] = []

    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not isinstance(a, str) or not a.strip():
            continue

        a_str = a.strip()

        if a_str.upper().startswith("STUDENT"):
            pending_student_rows.append(r)
            continue

        if a_str == "Preceptor":
            continue

        preceptor = format_preceptor(a_str)
        student_rows = pending_student_rows
        pending_student_rows = []

        for c, date in col_dates.items():
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue

            sh = parse_shift(v)
            if not sh:
                continue

            for sr in student_rows:
                raw = ws.cell(sr, c).value
                student = normalize_student(raw)
                if not student:
                    continue

                rows.append({
                    "Student Name": student,
                    "Date (YYYY-MM-DD)": date.isoformat(),
                    "Start Time (HH:MM)": sh["start"],
                    "End Time (HH:MM)": sh["end"],
                    "Location": "ACP",
                    "Station": sh["station"],
                    "Ambulance Number": sh["ambulance"],
                    "Preceptor": preceptor,
                })

    return pd.DataFrame(rows)


# ===================== PUBLIC API =====================

def extract_rows_from_workbook(xlsx_bytes: bytes, year: int, mode: str) -> pd.DataFrame:
    wb = load_workbook(filename=bytes_to_filelike(xlsx_bytes), data_only=True)
    if mode.upper() == "ACP":
        return extract_acp_rows(wb, year)
    return extract_pcp_rows(wb, year)


def apply_template_columns(extracted: pd.DataFrame, template_csv_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    tpl = pd.read_csv(template_csv_path)
    tpl_cols = list(tpl.columns)

    if extracted.empty:
        out = pd.DataFrame(columns=tpl_cols)
    else:
        out = extracted[tpl_cols].copy()

    debug = extracted.copy()
    return out, debug

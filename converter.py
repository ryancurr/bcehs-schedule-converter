import datetime as dt
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook


# ===================== SHARED RULES =====================

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

# =======================================================

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
        return f"0{x[0]}:{x[1:]}"
    return f"{x[:-2]}:{x[-2:]}"


def parse_shift(text: str) -> Optional[dict]:
    if not isinstance(text, str):
        return None

    t = text.replace("\n", " ").strip()

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
    s = re.sub(r"\s+", " ", name.strip())
    if "," in s:
        last, rest = s.split(",", 1)
        return f"{rest.strip()} {last.strip()}"
    return s


def format_preceptor(name: str) -> str:
    parts = [p.strip() for p in name.split("/") if p.strip()]
    return " / ".join(format_preceptor_one(p) for p in parts)


def is_partner(s: str) -> bool:
    s = s.lower().strip()
    return s in PARTNER_MARKERS or "partner" in s or "parnter" in s


def normalize_student(raw: str) -> str:
    if not isinstance(raw, str):
        return ""

    s = re.sub(r"\s+", " ", raw.strip())
    if not s:
        return ""

    if is_partner(s):
        return ""

    if not COLUMBIA_SUFFIX_PAT.search(s):
        return ""

    s = re.sub(COLUMBIA_SUFFIX_PAT, "", s).strip()

    if s.lower() in STUDENT_ALIASES:
        s = STUDENT_ALIASES[s.lower()]

    if s.lower() in EXCLUDE_STUDENTS:
        return ""

    return s


def parse_header_dates(ws, year: int, start_col: int) -> Dict[int, dt.date]:
    col_dates = {}
    for c in range(start_col, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str):
            m = HEADER_DATE_PAT.search(v)
            if m:
                mon = MONTHS[m.group(1).lower()]
                col_dates[c] = dt.date(year, mon, int(m.group(2)))
    return col_dates


# ===================== PCP =====================

def extract_pcp_rows(wb, year: int) -> pd.DataFrame:
    rows = []

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        col_dates = parse_header_dates(ws, year, 2)
        if not col_dates:
            continue

        for r in range(2, ws.max_row + 1):
            pre = ws.cell(r, 1).value
            if not isinstance(pre, str) or pre.strip().upper() in {"STUDENT", "PRECEPTOR"}:
                continue

            preceptor = format_preceptor_one(pre.strip())

            student_row = None
            if isinstance(ws.cell(r + 1, 1).value, str) and ws.cell(r + 1, 1).value.upper().startswith("STUDENT"):
                student_row = r + 1
            elif isinstance(ws.cell(r - 1, 1).value, str) and ws.cell(r - 1, 1).value.upper().startswith("STUDENT"):
                student_row = r - 1

            for c, date in col_dates.items():
                v = ws.cell(r, c).value
                if not isinstance(v, str):
                    continue

                sh = parse_shift(v)
                if not sh:
                    continue

                student_raw = ws.cell(student_row, c).value if student_row else ""
                student = normalize_student(student_raw)
                if not student:
                    continue

                rows.append({
                    "Student Name": student,
                    "Date (YYYY-MM-DD)": date.isoformat(),
                    "Start Time (HH:MM)": sh["start"],
                    "End Time (HH:MM)": sh["end"],
                    "Location": sheet,
                    "Station": sh["station"],
                    "Ambulance Number": sh["ambulance"],
                    "Preceptor": preceptor,
                })

    return pd.DataFrame(rows)


# ===================== ACP =====================

def extract_acp_rows(wb, year: int) -> pd.DataFrame:
    ws = wb[wb.sheetnames[0]]
    col_dates = parse_header_dates(ws, year, 3)

    rows = []
    pending_students = []

    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not isinstance(a, str):
            continue

        a = a.strip()

        if a.upper().startswith("STUDENT"):
            pending_students.append(r)
            continue

        if a.upper() == "PRECEPTOR":
            continue

        preceptor = format_preceptor(a)
        student_rows = pending_students
        pending_students = []

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
    return extract_acp_rows(wb, year) if mode.upper() == "ACP" else extract_pcp_rows(wb, year)


def apply_template_columns(df: pd.DataFrame, template_csv_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    tpl = pd.read_csv(template_csv_path)
    cols = list(tpl.columns)
    return df[cols].copy() if not df.empty else pd.DataFrame(columns=cols), df.copy()

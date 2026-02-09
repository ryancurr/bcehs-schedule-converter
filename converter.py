import datetime as dt
import re
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook


# ---- Your business rules ----
STUDENT_ALIASES = {
    "rory": "Rory-lynn Bradshaw",
}

PARTNER_MARKERS = {
    "partner", "parnter", "partnre", "parter", "prtnr",
}
# -----------------------------

TIME_RANGE_PAT = re.compile(
    r"(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})|(\d{3,4})\s*[-–]\s*(\d{3,4})"
)
CODE_TOKEN_PAT = re.compile(r"^\d{3}[A-Za-z0-9]+$")  # e.g., 240B1DA070
AMBULANCE_PAT = re.compile(r"^(\d{3}[A-Z]\d)", re.I)  # first 5 chars like 240B1
HEADER_DATE_PAT = re.compile(r"\b([A-Za-z]{3})/(\d{1,2})\b")

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


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
    if not text or not isinstance(text, str):
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


def format_preceptor(name: str) -> str:
    if not isinstance(name, str):
        return ""
    s = re.sub(r"\s+", " ", name.strip())
    if "," in s:
        last, rest = s.split(",", 1)
        return re.sub(r"\s+", " ", f"{rest.strip()} {last.strip()}".strip())
    return s


def clean_student(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = re.sub(r"\s+", " ", s.strip())

    # strip suffix like " - Columbia" or any " - Something"
    s = re.sub(r"\s*-\s*[A-Za-z][A-Za-z0-9 &/()'.-]+$", "", s).strip()

    key = s.lower()
    if key in STUDENT_ALIASES:
        return STUDENT_ALIASES[key]
    return s


def is_partner_marker(student_text: str) -> bool:
    s = (student_text or "").strip().lower()
    if s in PARTNER_MARKERS:
        return True
    # catch embedded variants
    if "partner" in s or "parnter" in s:
        return True
    return False


def is_group_header(text: str) -> bool:
    if not isinstance(text, str):
        return False
    s = text.strip()
    if s.upper() == "STUDENT":
        return False
    if "," in s:
        return False
    if s.upper() == s and len(s) >= 3:
        return True
    if any(k in s for k in ["Metro", "Vancouver", "Fraser", "Interior", "Island", "&", "SEA TO SKY", "COASTAL"]):
        return len(s.split()) >= 2
    return False


def parse_header_dates(ws, year: int) -> Dict[int, dt.date]:
    col_dates: Dict[int, dt.date] = {}
    for c in range(2, ws.max_column + 1):
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


def extract_rows_from_workbook(xlsx_bytes: bytes, year: int) -> pd.DataFrame:
    wb = load_workbook(filename=bytes_to_filelike(xlsx_bytes), data_only=True)

    all_rows: List[dict] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        col_dates = parse_header_dates(ws, year)
        if not col_dates:
            continue

        current_group = None

        for r in range(2, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if not isinstance(a, str):
                continue

            a_str = a.strip()
            if not a_str:
                continue

            if a_str.upper() == "STUDENT" or a_str == "Preceptor":
                continue

            if is_group_header(a_str):
                current_group = a_str
                continue

            preceptor = format_preceptor(a_str)

            # find the STUDENT row adjacent
            student_row = None
            up = ws.cell(r - 1, 1).value
            dn = ws.cell(r + 1, 1).value
            if isinstance(up, str) and up.strip().upper() == "STUDENT":
                student_row = r - 1
            elif isinstance(dn, str) and dn.strip().upper() == "STUDENT":
                student_row = r + 1

            for c, date in col_dates.items():
                cell_val = ws.cell(r, c).value
                if not isinstance(cell_val, str):
                    continue

                txt = cell_val.strip()
                if not txt or txt == "\\":
                    continue

                sh = parse_shift(txt)
                if not sh:
                    continue

                student_raw = ""
                if student_row:
                    sv = ws.cell(student_row, c).value
                    if isinstance(sv, str):
                        student_raw = sv.strip()

                # your filtering rules
                if is_partner_marker(student_raw):
                    continue
                if student_raw.strip() == "" or student_raw.strip().lower() == "student":
                    continue

                student = clean_student(student_raw)
                if not student:
                    continue

                location = current_group if current_group else sheet_name
                if current_group and current_group != sheet_name:
                    location = f"{sheet_name} - {current_group}"

                all_rows.append(
                    {
                        "Student Name": student,
                        "Date (YYYY-MM-DD)": date.isoformat(),
                        "Start Time (HH:MM)": sh["start"],
                        "End Time (HH:MM)": sh["end"],
                        "Location": location,
                        "Station": sh["station"],
                        "Ambulance Number": sh["ambulance"],
                        "Preceptor": preceptor,
                        "_Raw Shift Text": sh["raw"],
                        "_Code": sh["code"],
                        "_Sheet": sheet_name,
                    }
                )

    return pd.DataFrame(all_rows)


def apply_template_columns(extracted: pd.DataFrame, template_csv_bytes: bytes) -> (pd.DataFrame, pd.DataFrame):
    tpl = pd.read_csv(bytes_to_filelike(template_csv_bytes))
    tpl_cols = list(tpl.columns)

    if extracted.empty:
        out = pd.DataFrame(columns=tpl_cols)
    else:
        out = extracted[tpl_cols].copy()

    debug = extracted.copy()
    return out, debug


# --- helper to let openpyxl/pandas read uploaded bytes cleanly ---
def bytes_to_filelike(b: bytes):
    import io
    return io.BytesIO(b)

import datetime as dt
import re
from typing import Dict, List, Optional, Tuple, Union

import pandas as pd
from openpyxl import load_workbook


# -------- Rules (apply to BOTH ACP + PCP) --------

# If BCEHS types a short/incorrect name:
STUDENT_ALIASES = {
    "rory": "Rory-lynn Bradshaw",
}

# Known non-students you never want exported (optional list)
EXCLUDE_STUDENTS = {
    "jadyn",
    "jadyn langley",
}

PARTNER_MARKERS = {
    "partner", "parnter", "partnre", "parter", "prtnr",
}

COLUMBIA_SUFFIX_PAT = re.compile(r"\s*-\s*Columbia\s*$", re.I)

# -----------------------------------------------

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

    if not (code or start or end):
        return None

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
    """
    ACP sometimes has "Wilson, Travis / Johnston, Heather"
    We convert each side to First Last and keep " / " between.
    """
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
    """
    Shared student rules:
    - Must end with "- Columbia" or gets excluded
    - Remove "- Columbia" before output
    - Apply aliases
    - Apply explicit excludes
    """
    if not isinstance(raw, str):
        return ""

    s = re.sub(r"\s+", " ", raw.strip())
    if not s:
        return ""

    low = s.lower().strip()
    if low in {"student", "n/a", "na"}:
        return ""
    if is_partner_marker(s):
        return ""

    # REQUIRED: must have Columbia suffix
    if not COLUMBIA_SUFFIX_PAT.search(s):
        return ""

    # remove suffix for output
    s = re.sub(r"\s*-\s*Columbia\s*$", "", s, flags=re.I).strip()

    # alias
    if s.lower() in STUDENT_ALIASES:
        s = STUDENT_ALIASES[s.lower()]

    # exclude list
    if s.lower() in EXCLUDE_STUDENTS:
        return ""

    return s


def is_group_header(text: str) -> bool:
    """
    Best-effort section header detection (Metro Vancouver..., Fraser Valley, etc.)
    """
    if not isinstance(text, str):
        return False
    s = text.strip()
    if s.upper().startswith("STUDENT"):
        return False
    if "," in s:
        return False
    if s.upper() == s and len(s) >= 3:
        return True
    if any(k in s for k in ["Metro", "Vancouver", "Fraser", "Interior", "Island", "&", "SEA TO SKY", "COASTAL"]):
        return len(s.split()) >= 2
    return False


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


# ---------------- PCP extractor ----------------
def extract_pcp_rows(wb, year: int) -> pd.DataFrame:
    """
    PCP file (.xlsx): multiple region sheets.
    Layout: dates start at column B, preceptor in col A, student row adjacent labeled 'STUDENT'.
    Student row applies to the preceptor row above or below (we check both).
    """
    all_rows: List[dict] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        col_dates = parse_header_dates(ws, year, start_col=2)  # B
        if not col_dates:
            continue

        current_group = None

        for r in range(2, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if not isinstance(a, str):
                continue
            a_str = a.strip()
            if not a_str or a_str == "Preceptor":
                continue
            if a_str.upper() == "STUDENT":
                continue

            if is_group_header(a_str):
                current_group = a_str
                continue

            preceptor = format_preceptor_one(a_str)

            # locate adjacent STUDENT row
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

                student = normalize_student(student_raw)
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
                        "_Sheet": sheet_name,
                        "_Code": sh["code"],
                        "_Raw Shift Text": sh["raw"],
                    }
                )

    return pd.DataFrame(all_rows)


# ---------------- ACP extractor ----------------
def is_student_marker(v) -> bool:
    return isinstance(v, str) and v.strip().upper().startswith("STUDENT")


def extract_acp_rows(wb, year: int) -> pd.DataFrame:
    """
    ACP file (.xlsm): single sheet.
    Layout: dates start at column C (A=preceptor, B=email).
    IMPORTANT: student rows (STUDENT 1 / STUDENT 2) typically apply to the NEXT preceptor row below.
    So we accumulate student rows until a preceptor row appears, then apply those students to that preceptor row.
    """
    ws = wb[wb.sheetnames[0]]
    col_dates = parse_header_dates(ws, year, start_col=3)  # C

    rows: List[dict] = []
    current_group = None
    pending_student_rows: List[int] = []

    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not isinstance(a, str) or not a.strip():
            continue

        a_str = a.strip()

        if is_group_header(a_str):
            current_group = a_str
            pending_student_rows = []
            continue

        if a_str == "Preceptor":
            continue

        if is_student_marker(a_str):
            pending_student_rows.append(r)
            continue

        # preceptor row
        preceptor = format_preceptor(a_str)
        student_rows = pending_student_rows
        pending_student_rows = []

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

            # gather up to 2 students (or more if BCEHS adds more student rows)
            collected: List[str] = []
            for sr in student_rows:
                sv = ws.cell(sr, c).value
                if isinstance(sv, str) and sv.strip():
                    student = normalize_student(sv)
                    if student:
                        collected.append(student)

            # de-dupe (in case BCEHS repeats)
            seen = set()
            students: List[str] = []
            for s in collected:
                key = s.lower()
                if key not in seen:
                    seen.add(key)
                    students.append(s)

            if not students:
                continue

            location = current_group if current_group else "ACP"

            for student in students:
                rows.append(
                    {
                        "Student Name": student,
                        "Date (YYYY-MM-DD)": date.isoformat(),
                        "Start Time (HH:MM)": sh["start"],
                        "End Time (HH:MM)": sh["end"],
                        "Location": location,
                        "Station": sh["station"],
                        "Ambulance Number": sh["ambulance"],
                        "Preceptor": preceptor,
                        "_Sheet": ws.title,
                        "_Code": sh["code"],
                        "_Raw Shift Text": sh["raw"],
                    }
                )

    return pd.DataFrame(rows)


# --------------- Public API used by Streamlit ---------------
def extract_rows_from_workbook(xlsx_bytes: bytes, year: int, mode: str) -> pd.DataFrame:
    """
    mode: "PCP" or "ACP"
    """
    wb = load_workbook(filename=bytes_to_filelike(xlsx_bytes), data_only=True)

    mode = (mode or "").strip().upper()
    if mode == "ACP":
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

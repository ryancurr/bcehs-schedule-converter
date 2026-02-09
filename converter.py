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
    "jan": 1, "feb

"""Phone-number -> employee/department lookup.

Many bills (T-Mobile business, internet, fleet) list charges by phone number,
not employee name. This module loads the Line List and lets the categorizer
map any phone number on a bill straight to a department/category.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Optional

import openpyxl


_DIGITS_RE = re.compile(r"\D+")


def normalize_phone(raw) -> Optional[str]:
    """Normalize any US phone format to 10 digits, e.g. '5162723275'.

    Accepts '(516) 272-3275', '516-272-3275', '5162723275', 5162723275 (int), etc.
    Returns None if the input can't produce a 10-digit number.
    """
    if raw is None:
        return None
    digits = _DIGITS_RE.sub("", str(raw))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) != 10:
        return None
    return digits


@dataclass
class PhoneLine:
    phone: str            # 10 digits
    name: str
    department: str
    device_type: str = ""
    line_type: str = ""


class PhoneMatcher:
    def __init__(self, xlsx_path: str) -> None:
        self.by_phone: Dict[str, PhoneLine] = {}
        self._load(xlsx_path)

    def _load(self, path: str) -> None:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        headers = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                headers = [str(c).strip() if c is not None else "" for c in row]
                continue
            if not row:
                continue
            rec = {headers[j]: row[j] for j in range(min(len(headers), len(row)))}
            phone = normalize_phone(rec.get("Line"))
            if not phone:
                continue
            name = str(rec.get("First/Last name The full name associated with the line") or "").strip()
            dept = str(rec.get("Department") or "").strip()
            device = str(rec.get("Device Type") or "").strip()
            ltype = str(rec.get("Type") or "").strip()
            # Normalize common typos in dept names.
            dept = _normalize_dept(dept)
            self.by_phone[phone] = PhoneLine(
                phone=phone,
                name=name,
                department=dept,
                device_type=device,
                line_type=ltype,
            )

    def lookup(self, raw_phone) -> Optional[PhoneLine]:
        p = normalize_phone(raw_phone)
        if not p:
            return None
        return self.by_phone.get(p)


_DEPT_FIXES = {
    "Serivce": "Service",  # typo in source data
}


def _normalize_dept(d: str) -> str:
    return _DEPT_FIXES.get(d, d)

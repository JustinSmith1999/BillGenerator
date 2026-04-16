"""Bill file parsers.

Supports PDF (pdfplumber), Excel/CSV (pandas), and scanned images (pytesseract).
Every parser returns a list[BillLine] with a best-effort (name, amount) pair.

The heuristic is intentionally conservative: we look for lines that contain
BOTH a person-like name AND a currency amount. Unrecognized lines are kept
as 'unmatched' so the user can review them in the UI.
"""
from __future__ import annotations

import csv
import os
import re
from dataclasses import dataclass, field
from typing import List, Optional

import pdfplumber


# Monetary amounts: $1,234.56  1234.56  1,234  (optionally with leading $)
_AMOUNT_RE = re.compile(r"\$?\s*([\-\(]?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?\)?)")
# A "name-ish" token: two or three capitalized words, or "Last, First".
_NAME_RE = re.compile(
    r"(?:[A-Z][a-zA-Z\-']+\s*,\s*[A-Z][a-zA-Z\-']+(?:\s+[A-Z][a-zA-Z\-']+)?)"
    r"|(?:[A-Z][a-zA-Z\-']+(?:\s+[A-Z][a-zA-Z\-']+){1,2})"
)
# US phone number — require a recognizable format to avoid matching random
# 10-digit sequences like promo IDs ("2025022511055"). We accept:
#   (NNN) NNN-NNNN      canonical bill format
#   NNN-NNN-NNNN        dashed
#   NNN.NNN.NNNN        dotted
# Area code must start with 2-9 per NANP.
_PHONE_RE = re.compile(
    r"(?:\(\s*[2-9]\d{2}\s*\)\s*\d{3}[\s\-.]\d{4})"
    r"|(?:\b[2-9]\d{2}[\s\-.]\d{3}[\s\-.]\d{4}\b)"
)
# Line that looks like a T-Mobile per-phone summary row. Example:
#   "(516) 272-3275 Sunation Solar Syste ms p.55 $40.00 - - $8.00 -$30.00 $7.50 $25.50"
# We pick the LAST dollar amount as the line total.
_ALL_AMOUNTS_RE = re.compile(r"-?\$?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})")


@dataclass
class BillLine:
    raw: str
    name: Optional[str] = None
    amount: Optional[float] = None
    phone: Optional[str] = None      # 10-digit normalized, if any
    source_row: int = -1


def _parse_amount(text: str) -> Optional[float]:
    m = _AMOUNT_RE.search(text)
    if not m:
        return None
    raw = m.group(1).replace(",", "").replace("(", "-").replace(")", "")
    try:
        return float(raw)
    except ValueError:
        return None


def _extract_name(text: str) -> Optional[str]:
    m = _NAME_RE.search(text)
    return m.group(0).strip() if m else None


def _extract_phone(text: str) -> Optional[str]:
    m = _PHONE_RE.search(text)
    if not m:
        return None
    digits = re.sub(r"\D", "", m.group(0))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else None


def _extract_last_amount(text: str) -> Optional[float]:
    """For T-Mobile-style rows, the line total is the LAST dollar figure."""
    matches = _ALL_AMOUNTS_RE.findall(text)
    if not matches:
        return None
    raw = matches[-1].replace("$", "").replace(",", "").replace(" ", "").replace("(", "-").replace(")", "")
    try:
        return float(raw)
    except ValueError:
        return None


def _parse_text_lines(lines: List[str]) -> List[BillLine]:
    out: List[BillLine] = []
    seen_phones: set = set()  # dedupe — a phone appears in both summary and detail
    for i, raw in enumerate(lines):
        raw = raw.strip()
        if not raw:
            continue
        phone = _extract_phone(raw)
        # For phone-bearing lines, keep only the FIRST occurrence of each phone
        # (that's the per-line total on the summary pages). Later pages break
        # the same total down into components; summing them would double-count.
        if phone:
            if phone in seen_phones:
                continue
            seen_phones.add(phone)
            amount = _extract_last_amount(raw)
            name = _extract_name(raw)
            out.append(BillLine(raw=raw, name=name, amount=amount, phone=phone, source_row=i))
            continue
        # Non-phone line — use the older name+amount heuristic.
        name = _extract_name(raw)
        amount = _parse_amount(raw)
        if name or amount is not None:
            out.append(BillLine(raw=raw, name=name, amount=amount, source_row=i))
    return out


# ---- PDF ----------------------------------------------------------------

def parse_pdf(path: str) -> List[BillLine]:
    lines: List[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            # Try tabular extraction first.
            tables = page.extract_tables() or []
            for table in tables:
                for row in table:
                    lines.append(" | ".join(str(c) if c is not None else "" for c in row))
            # Always add free text too (many bills mix prose + tables).
            text = page.extract_text() or ""
            for line in text.splitlines():
                lines.append(line)
    return _parse_text_lines(lines)


# ---- Excel / CSV --------------------------------------------------------

def parse_spreadsheet(path: str) -> List[BillLine]:
    ext = os.path.splitext(path)[1].lower()
    rows: List[List[str]] = []
    if ext in (".csv", ".tsv"):
        delim = "\t" if ext == ".tsv" else ","
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f, delimiter=delim)
            for r in reader:
                rows.append([str(c) for c in r])
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                rows.append([str(c) if c is not None else "" for c in r])

    out: List[BillLine] = []
    # Detect a 'name' column and a numeric 'amount' column if a header exists.
    header = rows[0] if rows else []
    header_lower = [h.lower() for h in header]
    name_idx = next((i for i, h in enumerate(header_lower)
                     if any(k in h for k in ("employee", "name", "tech", "driver"))), -1)
    amount_idx = next((i for i, h in enumerate(header_lower)
                       if any(k in h for k in ("amount", "total", "cost", "charge", "price"))), -1)

    start = 1 if name_idx >= 0 or amount_idx >= 0 else 0
    for i, row in enumerate(rows[start:], start=start):
        if not any(row):
            continue
        joined = " | ".join(row)
        name = None
        amount = None
        if name_idx >= 0 and name_idx < len(row):
            name = row[name_idx].strip() or None
        if amount_idx >= 0 and amount_idx < len(row):
            amount = _parse_amount(row[amount_idx])
        if name is None:
            name = _extract_name(joined)
        if amount is None:
            amount = _parse_amount(joined)
        out.append(BillLine(raw=joined, name=name, amount=amount, source_row=i))
    return out


# ---- Image (OCR) --------------------------------------------------------

def parse_image(path: str) -> List[BillLine]:
    try:
        import pytesseract
        from PIL import Image
    except ImportError as e:
        raise RuntimeError(
            "OCR requires pytesseract + Pillow. Install them and the Tesseract binary."
        ) from e
    img = Image.open(path)
    text = pytesseract.image_to_string(img)
    return _parse_text_lines(text.splitlines())


# ---- Dispatcher ---------------------------------------------------------

def parse_bill(path: str) -> List[BillLine]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return parse_pdf(path)
    if ext in (".xlsx", ".xls", ".xlsm", ".csv", ".tsv"):
        return parse_spreadsheet(path)
    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"):
        return parse_image(path)
    raise ValueError(f"Unsupported bill file type: {ext}")

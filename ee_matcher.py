"""Employee list loader and fuzzy name matcher.

Accepts multiple name formats on bills ("First Last", "Last, First", "F. Last")
and normalizes them before matching to the Salesforce / EE list.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

import openpyxl
from rapidfuzz import fuzz, process


_PUNCT_RE = re.compile(r"[^\w\s]")


def _normalize(s: str) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = s.encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip()
    s = _PUNCT_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s)
    return s


def _split_last_first(raw: str) -> Tuple[str, str]:
    """Given a name like 'Smith, John' or 'John Smith', return (first, last)."""
    raw = raw.strip()
    if "," in raw:
        last, _, rest = raw.partition(",")
        return rest.strip(), last.strip()
    parts = raw.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def _name_variants(raw: str) -> List[str]:
    """Produce normalized variants of a name to improve fuzzy matching."""
    first, last = _split_last_first(raw)
    variants = {
        _normalize(raw),
        _normalize(f"{first} {last}"),
        _normalize(f"{last} {first}"),
        _normalize(f"{last}, {first}"),
    }
    return [v for v in variants if v]


@dataclass
class Employee:
    name: str
    department: str
    variants: List[str]


class EEMatcher:
    def __init__(
        self,
        xlsx_path: str,
        name_col: str = "HR: Employee Name",
        dept_col: str = "Department",
        threshold: int = 85,
    ) -> None:
        self.xlsx_path = xlsx_path
        self.name_col = name_col
        self.dept_col = dept_col
        self.threshold = threshold
        self.employees: List[Employee] = []
        # variant-string -> Employee (for fast exact lookup)
        self._variant_index: Dict[str, Employee] = {}
        self._load()

    def _load(self) -> None:
        wb = openpyxl.load_workbook(self.xlsx_path, data_only=True)
        ws = wb.active
        headers: List[str] = []
        for row_i, row in enumerate(ws.iter_rows(values_only=True)):
            if row_i == 0:
                headers = [str(c) if c is not None else "" for c in row]
                continue
            if not row or all(c is None for c in row):
                continue
            rec = {headers[i]: row[i] for i in range(len(headers))}
            name = rec.get(self.name_col)
            dept = rec.get(self.dept_col)
            if not name:
                continue
            name_s = str(name).strip()
            dept_s = str(dept).strip() if dept else ""
            emp = Employee(name=name_s, department=dept_s, variants=_name_variants(name_s))
            self.employees.append(emp)
            for v in emp.variants:
                self._variant_index.setdefault(v, emp)

    def update_from_salesforce(self, sf_rows: Iterable[Tuple[str, str]]) -> int:
        """Merge live Salesforce (name, department) pairs on top of the local list.

        Salesforce wins on conflicts (it's the live source of truth). Returns the
        number of employees added or updated.
        """
        changes = 0
        existing_by_name = {e.name: e for e in self.employees}
        for name, dept in sf_rows:
            if not name:
                continue
            name_s = str(name).strip()
            dept_s = str(dept).strip() if dept else ""
            emp = existing_by_name.get(name_s)
            if emp is None:
                emp = Employee(
                    name=name_s,
                    department=dept_s,
                    variants=_name_variants(name_s),
                )
                self.employees.append(emp)
                existing_by_name[name_s] = emp
                for v in emp.variants:
                    self._variant_index.setdefault(v, emp)
                changes += 1
            else:
                if dept_s and emp.department != dept_s:
                    emp.department = dept_s
                    changes += 1
        return changes

    def match(self, raw_name: str) -> Tuple[Optional[Employee], int]:
        """Return (best-match Employee, score 0-100). None if below threshold."""
        if not raw_name:
            return None, 0
        for v in _name_variants(raw_name):
            emp = self._variant_index.get(v)
            if emp is not None:
                return emp, 100
        # Fuzzy fallback against all known variants.
        choices = list(self._variant_index.keys())
        query = _normalize(raw_name)
        if not query or not choices:
            return None, 0
        best = process.extractOne(query, choices, scorer=fuzz.WRatio)
        if best is None:
            return None, 0
        variant, score, _ = best
        if score < self.threshold:
            return None, int(score)
        return self._variant_index[variant], int(score)

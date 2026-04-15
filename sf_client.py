"""Salesforce client wrapper.

Uses simple_salesforce for login and calls the Analytics/Reports REST endpoint
to pull the live report. Output: a list of (employee_name, department) tuples.

The user's report (00OUX000006He412AC) has two columns:
    HR: Employee Name | Department
but the column API names vary by org. We inspect the report metadata and pick
the columns that look like a name and a department.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Iterable, List, Optional, Tuple

try:
    from simple_salesforce import Salesforce
except ImportError:  # pragma: no cover
    Salesforce = None  # surfaced at runtime with a clearer message


NAME_HINTS = ("employee name", "full name", "name", "hr: employee")
DEPT_HINTS = ("department", "dept")


def _pick_column(columns_meta: dict, hints: Iterable[str]) -> Optional[str]:
    """Find the first column whose label contains any hint (case-insensitive)."""
    for api_name, meta in columns_meta.items():
        label = str(meta.get("label", "")).lower()
        if any(h in label for h in hints):
            return api_name
    return None


@dataclass
class SalesforceConfig:
    instance_url: str
    username: str
    password: str
    security_token: str
    domain: str
    report_id: str
    api_version: str


class SFClient:
    def __init__(self, cfg: SalesforceConfig) -> None:
        if Salesforce is None:
            raise RuntimeError(
                "simple_salesforce is not installed. Run: pip install simple-salesforce"
            )
        self.cfg = cfg
        # domain="login" for production, "test" for sandbox.
        self.sf = Salesforce(
            username=cfg.username,
            password=cfg.password,
            security_token=cfg.security_token,
            domain=cfg.domain or "login",
            version=cfg.api_version,
        )

    def fetch_report_rows(self) -> List[Tuple[str, str]]:
        """Run the live report and return (name, department) rows."""
        endpoint = f"analytics/reports/{self.cfg.report_id}?includeDetails=true"
        raw = self.sf.restful(endpoint, method="GET")
        return _parse_report(raw)


def _parse_report(raw: dict) -> List[Tuple[str, str]]:
    meta = raw.get("reportMetadata", {}) or {}
    ext = raw.get("reportExtendedMetadata", {}) or {}
    detail_columns: List[str] = meta.get("detailColumns", []) or []
    columns_meta = ext.get("detailColumnInfo", {}) or {}

    name_col = _pick_column(columns_meta, NAME_HINTS)
    dept_col = _pick_column(columns_meta, DEPT_HINTS)

    # Fall back to positional (first = name, second = department) if we
    # couldn't identify them by label.
    if not name_col and detail_columns:
        name_col = detail_columns[0]
    if not dept_col and len(detail_columns) > 1:
        dept_col = detail_columns[1]

    if not name_col:
        raise RuntimeError("Could not locate a name column in the Salesforce report.")

    try:
        name_idx = detail_columns.index(name_col)
    except ValueError:
        name_idx = 0
    try:
        dept_idx = detail_columns.index(dept_col) if dept_col else -1
    except ValueError:
        dept_idx = -1

    rows: List[Tuple[str, str]] = []
    fact_map = raw.get("factMap", {}) or {}
    for key, bucket in fact_map.items():
        # Only the "T!T" (grand total) bucket contains unaggregated rows for
        # tabular reports; summary reports use group keys like "0!T".
        for r in bucket.get("rows", []) or []:
            cells = r.get("dataCells", []) or []
            if name_idx >= len(cells):
                continue
            name = cells[name_idx].get("label") or cells[name_idx].get("value") or ""
            dept = ""
            if 0 <= dept_idx < len(cells):
                dept = cells[dept_idx].get("label") or cells[dept_idx].get("value") or ""
            name = str(name).strip()
            dept = str(dept).strip()
            if name:
                rows.append((name, dept))
    return rows


def load_config(path: str) -> SalesforceConfig:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    sf = data.get("salesforce", {})
    return SalesforceConfig(
        instance_url=sf.get("instance_url", ""),
        username=sf.get("username", ""),
        password=sf.get("password", ""),
        security_token=sf.get("security_token", ""),
        domain=sf.get("domain", "login"),
        report_id=sf.get("report_id", ""),
        api_version=sf.get("api_version", "60.0"),
    )

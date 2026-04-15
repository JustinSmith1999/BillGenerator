"""Department -> category mapping."""
from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Dict, List, Optional


UNCATEGORIZED = "Uncategorized"


@dataclass
class CategoryMap:
    categories: List[str]
    department_to_category: Dict[str, str]

    @classmethod
    def load(cls, path: str) -> "CategoryMap":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls(
            categories=list(data.get("categories", [])),
            department_to_category={
                str(k).strip(): str(v).strip()
                for k, v in data.get("department_to_category", {}).items()
            },
        )

    def categorize(self, department: Optional[str]) -> str:
        if not department:
            return UNCATEGORIZED
        return self.department_to_category.get(department.strip(), UNCATEGORIZED)

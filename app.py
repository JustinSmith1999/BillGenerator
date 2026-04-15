"""BillCategorizer — desktop GUI.

Workflow:
  1. Loads employee list from the configured XLSX.
  2. (Optional) 'Refresh from Salesforce' button pulls the live report and
     merges name/department changes onto the local EE list.
  3. User drops (or browses to) a bill file: PDF, XLSX/CSV, or image.
  4. Each line is parsed for (name, amount); the name is matched against the
     EE list; the employee's department is mapped to one of the 5 categories.
  5. On-screen summary shows totals per category + unmatched rows.
  6. 'Export Excel' / 'Export PDF' buttons write a formatted report.

Configuration lives in config.json and category_map.json — editable at any
time, no rebuild needed.
"""
from __future__ import annotations

import json
import os
import sys
import threading
import traceback
from datetime import datetime
from tkinter import (
    END,
    Tk,
    filedialog,
    messagebox,
    ttk,
    StringVar,
)
from tkinter import Text as TkText
from tkinter import PhotoImage, Label as TkLabel

try:
    from PIL import Image, ImageTk  # Pillow gives us PNG resizing support
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

from bill_parser import BillLine, parse_bill
from categorizer import CategoryMap, UNCATEGORIZED
from ee_matcher import EEMatcher
from exporters import CategorizedBill, LineResult, export_excel, export_pdf

try:
    # tkinterdnd2 gives us real drag-and-drop on Windows.
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_AVAILABLE = True
except ImportError:
    TkinterDnD = None
    DND_FILES = None
    _DND_AVAILABLE = False


APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
CATEGORY_MAP_PATH = os.path.join(APP_DIR, "category_map.json")


def _resource_path(rel: str) -> str:
    """Resolve a file next to the app/exe — works in both dev and PyInstaller."""
    base = getattr(sys, "_MEIPASS", APP_DIR)
    candidate = os.path.join(base, rel)
    if os.path.exists(candidate):
        return candidate
    return os.path.join(APP_DIR, rel)


def load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


class App:
    def __init__(self, root) -> None:
        self.root = root
        self.root.title("Bill Categorizer")
        self.root.geometry("1100x720")

        # Load config and data.
        self.config = load_json(CONFIG_PATH)
        self.category_map = CategoryMap.load(CATEGORY_MAP_PATH)
        ee_path = self._resolve_path(self.config.get("employee_list_path", "data/employees.xlsx"))
        self.ee = EEMatcher(
            ee_path,
            name_col=self.config.get("employee_name_column", "HR: Employee Name"),
            dept_col=self.config.get("employee_department_column", "Department"),
            threshold=int(self.config.get("fuzzy_match_threshold", 85)),
        )

        self.current_bill: CategorizedBill | None = None
        self.status_var = StringVar(value=f"Loaded {len(self.ee.employees)} employees from local list.")
        self.bill_path_var = StringVar(value="")

        self._build_ui()

    def _resolve_path(self, p: str) -> str:
        if os.path.isabs(p):
            return p
        return os.path.join(APP_DIR, p)

    def _resolve_output_dir(self) -> str:
        """Pick a writable output folder.

        Preference order: config.output_folder (if set) -> Documents/BillCategorizer/reports.
        Documents is a safe default: it's writable from USB-run apps and it's a
        familiar place for non-technical users to find their exports.
        """
        override = (self.config.get("output_folder") or "").strip()
        if override:
            return self._resolve_path(override)
        docs = os.path.join(os.path.expanduser("~"), "Documents", "BillCategorizer", "reports")
        return docs

    # ---- UI -----------------------------------------------------------

    def _build_ui(self) -> None:
        # --- Branded header bar ---
        header = TkLabel(self.root, bg="#FFFFFF", bd=0)
        header.pack(fill="x")
        logo_path = _resource_path(os.path.join("assets", "logo.png"))
        self._logo_image = None  # keep reference so Tk doesn't garbage-collect
        if os.path.exists(logo_path):
            try:
                if _PIL_AVAILABLE:
                    img = Image.open(logo_path).convert("RGBA")
                    # scale down to ~60px tall while preserving aspect ratio
                    target_h = 60
                    ratio = target_h / img.height
                    img = img.resize(
                        (max(1, int(img.width * ratio)), target_h),
                        Image.LANCZOS,
                    )
                    self._logo_image = ImageTk.PhotoImage(img)
                else:
                    self._logo_image = PhotoImage(file=logo_path)
            except Exception:
                self._logo_image = None
        if self._logo_image is not None:
            logo_lbl = TkLabel(header, image=self._logo_image, bg="#FFFFFF")
            logo_lbl.pack(side="left", padx=12, pady=8)
        title_lbl = TkLabel(
            header,
            text="Bill Categorizer",
            font=("Segoe UI", 16, "bold"),
            fg="#1441A0",
            bg="#FFFFFF",
        )
        title_lbl.pack(side="left", padx=10)
        ttk.Separator(self.root, orient="horizontal").pack(fill="x")

        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Refresh from Salesforce", command=self.on_refresh_sf).pack(side="left")
        ttk.Button(top, text="Open Bill...", command=self.on_open_bill).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Re-run Categorization", command=self.on_rerun).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Export Excel", command=self.on_export_excel).pack(side="right")
        ttk.Button(top, text="Export PDF", command=self.on_export_pdf).pack(side="right", padx=(0, 8))

        # Drop zone
        drop = ttk.LabelFrame(self.root, text="Bill file", padding=10)
        drop.pack(fill="x", padx=10, pady=(0, 6))
        self.drop_label = ttk.Label(
            drop,
            textvariable=self.bill_path_var,
            anchor="w",
            relief="groove",
            padding=10,
        )
        self.drop_label.pack(fill="x")
        if _DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_dnd_drop)
            self.bill_path_var.set("Drag-and-drop a bill here, or click 'Open Bill...'.")
        else:
            self.bill_path_var.set("Click 'Open Bill...' to select a bill file.")

        # Summary panel
        mid = ttk.Frame(self.root, padding=10)
        mid.pack(fill="both", expand=True)

        left = ttk.LabelFrame(mid, text="Totals by category", padding=8)
        left.pack(side="left", fill="y")
        self.totals_tree = ttk.Treeview(left, columns=("total",), show="tree headings", height=10)
        self.totals_tree.heading("#0", text="Category")
        self.totals_tree.heading("total", text="Total")
        self.totals_tree.column("#0", width=180)
        self.totals_tree.column("total", width=120, anchor="e")
        self.totals_tree.pack(fill="both", expand=True)

        right = ttk.LabelFrame(mid, text="Line items", padding=8)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))
        cols = ("employee", "dept", "category", "amount", "score")
        self.lines_tree = ttk.Treeview(right, columns=cols, show="headings")
        for c, w in zip(cols, (200, 180, 120, 100, 60)):
            self.lines_tree.heading(c, text=c.title())
            self.lines_tree.column(c, width=w, anchor="w" if c not in ("amount", "score") else "e")
        self.lines_tree.pack(fill="both", expand=True)

        # Unmatched panel
        bottom = ttk.LabelFrame(self.root, text="Unmatched lines (review manually)", padding=8)
        bottom.pack(fill="x", padx=10, pady=(0, 6))
        self.unmatched_text = TkText(bottom, height=6, wrap="none")
        self.unmatched_text.pack(fill="x")

        # Status bar
        status = ttk.Frame(self.root, padding=(10, 4))
        status.pack(fill="x")
        ttk.Label(status, textvariable=self.status_var, anchor="w").pack(fill="x")

    def _on_dnd_drop(self, event) -> None:  # pragma: no cover - GUI only
        path = event.data.strip().strip("{").strip("}")
        if path:
            self.load_bill(path)

    # ---- Handlers -----------------------------------------------------

    def on_refresh_sf(self) -> None:
        self._set_status("Contacting Salesforce...")
        threading.Thread(target=self._refresh_sf_worker, daemon=True).start()

    def _refresh_sf_worker(self) -> None:
        try:
            from sf_client import SFClient, load_config
            cfg = load_config(CONFIG_PATH)
            if not cfg.username or not cfg.password:
                raise RuntimeError(
                    "Salesforce credentials are empty. Edit config.json first."
                )
            client = SFClient(cfg)
            rows = client.fetch_report_rows()
            changed = self.ee.update_from_salesforce(rows)
            self._set_status(
                f"Salesforce OK — {len(rows)} rows pulled, {changed} local records added/updated "
                f"({datetime.now():%H:%M:%S})."
            )
            # If a bill is already loaded, re-categorize with the fresh data.
            if self.current_bill is not None:
                self.on_rerun()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Salesforce error", str(e))
            self._set_status(f"Salesforce refresh failed: {e}")

    def on_open_bill(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a bill",
            filetypes=[
                ("Supported", "*.pdf *.xlsx *.xls *.xlsm *.csv *.tsv *.png *.jpg *.jpeg *.tif *.tiff *.bmp"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.load_bill(path)

    def load_bill(self, path: str) -> None:
        self.bill_path_var.set(path)
        self._set_status(f"Parsing bill: {os.path.basename(path)}")
        threading.Thread(target=self._parse_worker, args=(path,), daemon=True).start()

    def _parse_worker(self, path: str) -> None:
        try:
            lines = parse_bill(path)
            bill = self._categorize(os.path.basename(path), lines)
            self.current_bill = bill
            self.root.after(0, self._render_bill)
            matched = sum(1 for ln in bill.lines if ln.matched_name)
            self._set_status(
                f"Parsed {len(bill.lines)} lines — {matched} matched, "
                f"{len(bill.lines) - matched} unmatched. Grand total ${bill.grand_total():,.2f}."
            )
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Parse error", str(e))
            self._set_status(f"Parse failed: {e}")

    def on_rerun(self) -> None:
        if self.current_bill is None:
            return
        # Re-parse the raw lines (the source file may have changed or we may
        # have pulled new SF data).
        raw_lines = [
            BillLine(raw=ln.raw, name=None, amount=ln.amount)
            for ln in self.current_bill.lines
        ]
        bill = self._categorize(self.current_bill.bill_name, raw_lines)
        self.current_bill = bill
        self._render_bill()

    def _categorize(self, bill_name: str, lines) -> CategorizedBill:
        # Accept either BillLine or LineResult-ish objects.
        results: list[LineResult] = []
        for ln in lines:
            name_raw = getattr(ln, "name", None)
            amount = getattr(ln, "amount", None)
            raw = getattr(ln, "raw", "")
            emp = None
            score = 0
            if name_raw:
                emp, score = self.ee.match(name_raw)
            matched_name = emp.name if emp else None
            dept = emp.department if emp else None
            category = self.category_map.categorize(dept) if dept else UNCATEGORIZED
            results.append(
                LineResult(
                    raw=raw,
                    matched_name=matched_name,
                    department=dept,
                    category=category,
                    amount=amount,
                    score=score,
                )
            )
        return CategorizedBill(
            bill_name=bill_name,
            lines=results,
            categories=list(self.category_map.categories),
        )

    def _render_bill(self) -> None:
        self.totals_tree.delete(*self.totals_tree.get_children())
        self.lines_tree.delete(*self.lines_tree.get_children())
        self.unmatched_text.delete("1.0", END)

        if self.current_bill is None:
            return

        totals = self.current_bill.totals()
        ordered = list(self.current_bill.categories) + [
            k for k in totals if k not in self.current_bill.categories
        ]
        for cat in ordered:
            self.totals_tree.insert(
                "", "end", text=cat, values=(f"${totals.get(cat, 0.0):,.2f}",)
            )
        self.totals_tree.insert(
            "", "end", text="— Grand total —",
            values=(f"${self.current_bill.grand_total():,.2f}",),
        )

        for ln in self.current_bill.lines:
            self.lines_tree.insert(
                "", "end",
                values=(
                    ln.matched_name or "—",
                    ln.department or "—",
                    ln.category,
                    f"${ln.amount:,.2f}" if ln.amount is not None else "—",
                    ln.score if ln.score else "",
                ),
            )

        unmatched = self.current_bill.unmatched()
        for ln in unmatched:
            amt = f"${ln.amount:,.2f}" if ln.amount is not None else "—"
            self.unmatched_text.insert(END, f"[{amt}]  {ln.raw[:200]}\n")

    def on_export_excel(self) -> None:
        if self.current_bill is None:
            messagebox.showinfo("Nothing to export", "Load a bill first.")
            return
        out_dir = self._resolve_output_dir()
        os.makedirs(out_dir, exist_ok=True)
        default = f"bill_breakdown_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        path = filedialog.asksaveasfilename(
            initialdir=out_dir, initialfile=default, defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        export_excel(self.current_bill, path)
        self._set_status(f"Excel written: {path}")
        messagebox.showinfo("Exported", f"Saved to:\n{path}")

    def on_export_pdf(self) -> None:
        if self.current_bill is None:
            messagebox.showinfo("Nothing to export", "Load a bill first.")
            return
        out_dir = self._resolve_output_dir()
        os.makedirs(out_dir, exist_ok=True)
        default = f"bill_breakdown_{datetime.now():%Y%m%d_%H%M%S}.pdf"
        path = filedialog.asksaveasfilename(
            initialdir=out_dir, initialfile=default, defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return
        export_pdf(self.current_bill, path)
        self._set_status(f"PDF written: {path}")
        messagebox.showinfo("Exported", f"Saved to:\n{path}")

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_var.set(text))


def main() -> None:
    if _DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

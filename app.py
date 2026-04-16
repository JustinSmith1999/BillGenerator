"""BillCategorizer — desktop GUI (modern branded UI).

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
import re as _re
import sys
import threading
import traceback
from datetime import datetime
from tkinter import (
    END,
    Canvas,
    Frame as TkFrame,
    Tk,
    filedialog,
    messagebox,
    ttk,
    StringVar,
)
from tkinter import Text as TkText
from tkinter import PhotoImage, Label as TkLabel, Button as TkButton
from typing import Optional

try:
    from PIL import Image, ImageTk
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

from bill_parser import BillLine, parse_bill
from categorizer import CategoryMap, UNCATEGORIZED
from ee_matcher import EEMatcher
from phone_matcher import PhoneMatcher
from exporters import CategorizedBill, LineResult, export_excel, export_pdf, export_ap_analysis

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_AVAILABLE = True
except ImportError:
    TkinterDnD = None
    DND_FILES = None
    _DND_AVAILABLE = False


APP_DIR = os.path.dirname(os.path.abspath(__file__))

_HOTSPOT_RE = _re.compile(r"(?i)hot\s*spot\s*#?\s*(\d+)")


def _hotspot_override(name: str, default_category: str) -> str:
    """HotSpot 1-10 -> Commercial, HotSpot 11+ -> Residential."""
    m = _HOTSPOT_RE.search(name)
    if not m:
        return default_category
    num = int(m.group(1))
    return "Commercial" if 1 <= num <= 10 else "Residential"


CONFIG_PATH = os.path.join(APP_DIR, "config.json")
CATEGORY_MAP_PATH = os.path.join(APP_DIR, "category_map.json")

# ---- Brand palette --------------------------------------------------------
_BLUE      = "#1441A0"
_CYAN      = "#1EADE6"
_YELLOW    = "#FFD200"
_WHITE     = "#FFFFFF"
_BG        = "#F4F6FA"      # light grey-blue page background
_CARD_BG   = "#FFFFFF"
_TEXT_DARK  = "#1E293B"
_TEXT_MED   = "#64748B"
_GREEN     = "#16A34A"
_RED       = "#DC2626"

# Category card colours (background tint, text)
_CAT_COLORS = {
    "Residential": ("#E0ECFF", _BLUE),
    "Commercial":  ("#DAFBE8", "#166534"),
    "Service":     ("#FEF3C7", "#92400E"),
    "Roofing":     ("#FEE2E2", "#991B1B"),
    "Executive":   ("#EDE9FE", "#5B21B6"),
}
_DEFAULT_CAT_COLOR = ("#F1F5F9", _TEXT_DARK)


def _resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", APP_DIR)
    candidate = os.path.join(base, rel)
    return candidate if os.path.exists(candidate) else os.path.join(APP_DIR, rel)


def load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Custom Tkinter widgets
# ---------------------------------------------------------------------------

class _RoundedCard(TkFrame):
    """A simple card-like frame (white bg, slight padding)."""
    def __init__(self, parent, **kw):
        kw.setdefault("bg", _CARD_BG)
        kw.setdefault("bd", 0)
        kw.setdefault("highlightthickness", 1)
        kw.setdefault("highlightbackground", "#E2E8F0")
        super().__init__(parent, **kw)


class _BrandButton(TkLabel):
    """Flat coloured button using a Label + hand cursor + click binding."""
    def __init__(self, parent, text="", command=None, bg=_BLUE, fg=_WHITE,
                 font=("Segoe UI", 10, "bold"), padx=18, pady=8, **kw):
        super().__init__(parent, text=text, bg=bg, fg=fg, font=font,
                         padx=padx, pady=pady, cursor="hand2", **kw)
        self._cmd = command
        self._base_bg = bg
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", lambda e: self.config(bg=self._hover_color()))
        self.bind("<Leave>", lambda e: self.config(bg=self._base_bg))

    def _hover_color(self):
        # Darken slightly
        try:
            r, g, b = self.winfo_rgb(self._base_bg)
            r, g, b = max(0, r // 256 - 18), max(0, g // 256 - 18), max(0, b // 256 - 18)
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return self._base_bg

    def _on_click(self, event=None):
        if self._cmd:
            self._cmd()


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root) -> None:
        self.root = root
        self.root.title("Sunation Energy  —  Bill Categorizer")
        self.root.geometry("1100x780")
        self.root.configure(bg=_BG)
        self.root.minsize(900, 600)

        # Load config and data
        self.config = load_json(CONFIG_PATH)
        self.category_map = CategoryMap.load(CATEGORY_MAP_PATH)
        ee_path = self._resolve_path(self.config.get("employee_list_path", "data/employees.xlsx"))
        self.ee = EEMatcher(
            ee_path,
            name_col=self.config.get("employee_name_column", "HR: Employee Name"),
            dept_col=self.config.get("employee_department_column", "Department"),
            threshold=int(self.config.get("fuzzy_match_threshold", 85)),
        )

        self.phones: PhoneMatcher | None = None
        phone_list_path = self.config.get("phone_list_path", "data/line_list.xlsx")
        if phone_list_path:
            full = self._resolve_path(phone_list_path)
            if os.path.exists(full):
                try:
                    self.phones = PhoneMatcher(full)
                except Exception as e:
                    print(f"Warning: could not load phone list: {e}")

        self.current_bill: CategorizedBill | None = None
        self.status_var = StringVar(value=f"Ready  —  {len(self.ee.employees)} employees loaded.")
        self.bill_path_var = StringVar(value="")
        self._progress_active = False
        self._category_labels: dict[str, TkLabel] = {}
        self._grand_total_label: TkLabel | None = None

        self._build_ui()

    # ---- path helpers -----------------------------------------------------

    def _resolve_path(self, p: str) -> str:
        return p if os.path.isabs(p) else os.path.join(APP_DIR, p)

    def _resolve_output_dir(self) -> str:
        override = (self.config.get("output_folder") or "").strip()
        if override:
            return self._resolve_path(override)
        return os.path.join(os.path.expanduser("~"), "Documents", "BillCategorizer", "reports")

    # ---- UI ---------------------------------------------------------------

    def _build_ui(self) -> None:
        # ===== Header bar ================================================
        header = TkFrame(self.root, bg=_WHITE, bd=0, highlightthickness=0)
        header.pack(fill="x")

        # Logo
        logo_path = _resource_path(os.path.join("assets", "logo.png"))
        self._logo_image = None
        if os.path.exists(logo_path):
            try:
                if _PIL_AVAILABLE:
                    img = Image.open(logo_path).convert("RGBA")
                    target_h = 48
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
            TkLabel(header, image=self._logo_image, bg=_WHITE).pack(
                side="left", padx=(16, 8), pady=10
            )

        TkLabel(
            header, text="Bill Categorizer", font=("Segoe UI", 18, "bold"),
            fg=_BLUE, bg=_WHITE,
        ).pack(side="left", padx=4, pady=10)

        # Header buttons (right side)
        btn_frame = TkFrame(header, bg=_WHITE)
        btn_frame.pack(side="right", padx=16, pady=10)

        _BrandButton(btn_frame, text="Open Bill", command=self.on_open_bill,
                      bg=_BLUE, fg=_WHITE).pack(side="left", padx=(0, 8))
        _BrandButton(btn_frame, text="Refresh Salesforce", command=self.on_refresh_sf,
                      bg=_CYAN, fg=_WHITE).pack(side="left", padx=(0, 8))
        _BrandButton(btn_frame, text="Export Excel", command=self.on_export_excel,
                      bg="#16A34A", fg=_WHITE).pack(side="left", padx=(0, 8))
        _BrandButton(btn_frame, text="Export PDF", command=self.on_export_pdf,
                      bg="#DC2626", fg=_WHITE).pack(side="left")

        # Accent stripe under header
        TkFrame(self.root, bg=_CYAN, height=3, bd=0).pack(fill="x")

        # ===== Progress bar (hidden until needed) =========================
        self._progress_frame = TkFrame(self.root, bg=_BG, height=6, bd=0)
        self._progress_frame.pack(fill="x")
        self._progress_bar = ttk.Progressbar(
            self._progress_frame, mode="indeterminate", length=400
        )
        # Not packed yet — shown dynamically

        # ===== Main scrollable area =======================================
        main = TkFrame(self.root, bg=_BG, bd=0)
        main.pack(fill="both", expand=True, padx=20, pady=(14, 6))

        # -- Drop zone / file label ----------------------------------------
        drop_card = _RoundedCard(main, padx=16, pady=10)
        drop_card.pack(fill="x", pady=(0, 14))

        self.drop_label = TkLabel(
            drop_card, textvariable=self.bill_path_var,
            bg=_CARD_BG, fg=_TEXT_MED, font=("Segoe UI", 10),
            anchor="w", padx=8, pady=6,
        )
        self.drop_label.pack(fill="x")
        if _DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_dnd_drop)
            self.bill_path_var.set("Drag-and-drop a bill here, or click  Open Bill")
        else:
            self.bill_path_var.set("Click  Open Bill  to select a bill file.")

        # -- Category totals cards -----------------------------------------
        self._cards_frame = TkFrame(main, bg=_BG, bd=0)
        self._cards_frame.pack(fill="x", pady=(0, 14))

        categories = list(self.category_map.categories)
        for cat in categories:
            bg_tint, fg = _CAT_COLORS.get(cat, _DEFAULT_CAT_COLOR)
            card = TkFrame(self._cards_frame, bg=bg_tint, bd=0,
                           highlightthickness=1, highlightbackground="#E2E8F0",
                           padx=16, pady=12)
            card.pack(side="left", fill="x", expand=True, padx=(0, 10))
            TkLabel(card, text=cat.upper(), font=("Segoe UI", 9, "bold"),
                    fg=fg, bg=bg_tint).pack(anchor="w")
            amt_lbl = TkLabel(card, text="$0.00", font=("Segoe UI", 22, "bold"),
                              fg=fg, bg=bg_tint)
            amt_lbl.pack(anchor="w", pady=(2, 0))
            self._category_labels[cat] = amt_lbl

        # Grand total card (full width, accent color)
        gt_card = TkFrame(self._cards_frame, bg=_BLUE, bd=0,
                          highlightthickness=0, padx=16, pady=12)
        gt_card.pack(side="left", fill="x", expand=True)
        TkLabel(gt_card, text="GRAND TOTAL", font=("Segoe UI", 9, "bold"),
                fg=_YELLOW, bg=_BLUE).pack(anchor="w")
        self._grand_total_label = TkLabel(
            gt_card, text="$0.00", font=("Segoe UI", 22, "bold"),
            fg=_WHITE, bg=_BLUE,
        )
        self._grand_total_label.pack(anchor="w", pady=(2, 0))
        self._overhead_label = TkLabel(
            gt_card, text="", font=("Segoe UI", 8),
            fg="#93C5FD", bg=_BLUE,
        )
        self._overhead_label.pack(anchor="w")

        # -- Line items table ----------------------------------------------
        table_card = _RoundedCard(main, padx=0, pady=0)
        table_card.pack(fill="both", expand=True, pady=(0, 8))

        table_header = TkFrame(table_card, bg=_CARD_BG, bd=0)
        table_header.pack(fill="x", padx=16, pady=(12, 4))
        TkLabel(table_header, text="Line Items", font=("Segoe UI", 12, "bold"),
                fg=_TEXT_DARK, bg=_CARD_BG).pack(side="left")
        self._line_count_lbl = TkLabel(
            table_header, text="", font=("Segoe UI", 10),
            fg=_TEXT_MED, bg=_CARD_BG,
        )
        self._line_count_lbl.pack(side="left", padx=(10, 0))

        # Treeview with simplified columns: Employee, Category, Amount
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Brand.Treeview",
                        background=_WHITE,
                        foreground=_TEXT_DARK,
                        fieldbackground=_WHITE,
                        font=("Segoe UI", 10),
                        rowheight=28)
        style.configure("Brand.Treeview.Heading",
                        background=_BLUE,
                        foreground=_WHITE,
                        font=("Segoe UI", 10, "bold"),
                        relief="flat")
        style.map("Brand.Treeview.Heading",
                  background=[("active", _CYAN)])
        style.map("Brand.Treeview",
                  background=[("selected", "#DBEAFE")],
                  foreground=[("selected", _BLUE)])

        cols = ("employee", "category", "amount")
        tree_frame = TkFrame(table_card, bg=_CARD_BG, bd=0)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.lines_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", style="Brand.Treeview",
        )
        for c, label, w, anc in [
            ("employee", "Employee", 340, "w"),
            ("category", "Category", 160, "w"),
            ("amount", "Amount", 130, "e"),
        ]:
            self.lines_tree.heading(c, text=label)
            self.lines_tree.column(c, width=w, anchor=anc, minwidth=80)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.lines_tree.yview)
        self.lines_tree.configure(yscrollcommand=vsb.set)
        self.lines_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Tag for alternating row colors
        self.lines_tree.tag_configure("oddrow", background="#F8FAFC")
        self.lines_tree.tag_configure("evenrow", background=_WHITE)
        self.lines_tree.tag_configure("unmatched", foreground=_RED)

        # ===== Status bar =================================================
        status_bar = TkFrame(self.root, bg=_WHITE, bd=0, highlightthickness=0)
        status_bar.pack(fill="x", side="bottom")
        TkFrame(status_bar, bg="#E2E8F0", height=1).pack(fill="x")
        TkLabel(status_bar, textvariable=self.status_var,
                font=("Segoe UI", 9), fg=_TEXT_MED, bg=_WHITE,
                anchor="w", padx=16, pady=6).pack(fill="x")

    # ---- Progress indicator -----------------------------------------------

    def _show_progress(self):
        self._progress_active = True
        self._progress_bar.pack(fill="x", padx=0, pady=0)
        self._progress_bar.start(12)

    def _hide_progress(self):
        self._progress_active = False
        self._progress_bar.stop()
        self._progress_bar.pack_forget()

    # ---- DnD --------------------------------------------------------------

    def _on_dnd_drop(self, event) -> None:
        path = event.data.strip().strip("{").strip("}")
        if path:
            self.load_bill(path)

    # ---- Handlers ---------------------------------------------------------

    def on_refresh_sf(self) -> None:
        self._set_status("Contacting Salesforce...")
        self.root.after(0, self._show_progress)
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
            self.root.after(0, self._hide_progress)
            self._set_status(
                f"Salesforce OK  —  {len(rows)} rows pulled, {changed} updated "
                f"({datetime.now():%H:%M:%S})"
            )
            if self.current_bill is not None:
                self.on_rerun()
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, self._hide_progress)
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
        self.bill_path_var.set(f"  {os.path.basename(path)}")
        self._set_status(f"Parsing bill: {os.path.basename(path)}  ...")
        self.root.after(0, self._show_progress)
        threading.Thread(target=self._parse_worker, args=(path,), daemon=True).start()

    def _parse_worker(self, path: str) -> None:
        try:
            lines = parse_bill(path)
            # If the bill has phone-bearing lines, use ONLY those.
            phone_lines = [ln for ln in lines if getattr(ln, "phone", None)]
            all_lines = lines  # keep full set to extract account-level charges
            if phone_lines:
                lines = phone_lines

            bill = self._categorize(os.path.basename(path), lines)

            # --- Grand total fix ---
            # If we filtered to phone-only lines, check for account-level
            # charges (the difference between the full-bill total and the
            # phone-line total). Distribute that overhead evenly.
            if phone_lines:
                full_total = sum(
                    getattr(ln, "amount", 0) or 0
                    for ln in all_lines
                    if getattr(ln, "amount", None) is not None
                )
                phone_total = bill.grand_total()
                # Try to extract the bill's stated total from page 1 text.
                stated_total = self._extract_stated_total(path)
                if stated_total is not None and stated_total > phone_total:
                    overhead = round(stated_total - phone_total, 2)
                elif full_total > phone_total * 1.01:
                    # fallback: use full parse total, but only if it's
                    # meaningfully more (>1% larger) to avoid noise
                    overhead = round(full_total - phone_total, 2)
                else:
                    overhead = 0.0

                if overhead > 0:
                    # Add a synthetic "Account-level charges" line
                    bill.lines.append(LineResult(
                        raw="Account-level charges (overhead)",
                        matched_name=None,
                        department=None,
                        category="Uncategorized",
                        amount=overhead,
                        score=0,
                    ))

            self.current_bill = bill
            self.root.after(0, self._hide_progress)
            self.root.after(0, self._render_bill)
            matched = sum(1 for ln in bill.lines if ln.matched_name)

            ap_path = self._auto_save_ap_analysis(bill)
            self._set_status(
                f"Done  —  {len(bill.lines)} lines, {matched} matched.  "
                f"Grand total ${bill.grand_total():,.2f}.  "
                f"AP analysis saved."
            )
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, self._hide_progress)
            messagebox.showerror("Parse error", str(e))
            self._set_status(f"Parse failed: {e}")

    def _extract_stated_total(self, path: str) -> Optional[float]:
        """Try to pull the bill's stated 'Total due' from page 1 of a PDF."""
        if not path.lower().endswith(".pdf"):
            return None
        try:
            import pdfplumber
            with pdfplumber.open(path) as pdf:
                if not pdf.pages:
                    return None
                text = pdf.pages[0].extract_text() or ""
                # Look for patterns like "Total due $4,786.27" or "Total: $4786.27"
                m = _re.search(
                    r"(?i)(?:total\s*(?:due|amount|charges?|balance)?)\s*[:\s]*\$?\s*([\d,]+\.\d{2})",
                    text,
                )
                if m:
                    return float(m.group(1).replace(",", ""))
        except Exception:
            pass
        return None

    def _auto_save_ap_analysis(self, bill: CategorizedBill) -> str:
        out_dir = self._resolve_output_dir()
        os.makedirs(out_dir, exist_ok=True)
        safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in bill.bill_name)
        safe_name = safe_name.replace(" ", "_")[:80]
        filename = f"AP_PhoneLine_Analysis_{safe_name}.xlsx"
        ap_path = os.path.join(out_dir, filename)
        export_ap_analysis(bill, ap_path)
        return ap_path

    def on_rerun(self) -> None:
        if self.current_bill is None:
            return
        raw_lines = [
            BillLine(raw=ln.raw, name=None, amount=ln.amount)
            for ln in self.current_bill.lines
        ]
        bill = self._categorize(self.current_bill.bill_name, raw_lines)
        self.current_bill = bill
        self._render_bill()

    def _categorize(self, bill_name: str, lines) -> CategorizedBill:
        results: list[LineResult] = []
        for ln in lines:
            name_raw = getattr(ln, "name", None)
            amount = getattr(ln, "amount", None)
            phone = getattr(ln, "phone", None)
            raw = getattr(ln, "raw", "")

            matched_name: Optional[str] = None
            dept: Optional[str] = None
            score = 0

            # Step 1: phone -> line list -> EE list
            if phone and self.phones is not None:
                phone_line = self.phones.lookup(phone)
                if phone_line:
                    emp, ee_score = self.ee.match(phone_line.name)
                    if emp:
                        matched_name = emp.name
                        dept = emp.department
                        score = max(90, ee_score)
                    else:
                        matched_name = phone_line.name
                        dept = phone_line.department or None
                        score = 60

            # Step 2: no phone match, try name-on-the-line
            if dept is None and name_raw:
                emp, ee_score = self.ee.match(name_raw)
                if emp:
                    matched_name = emp.name
                    dept = emp.department
                    score = ee_score

            category = self.category_map.categorize(dept) if dept else UNCATEGORIZED

            if matched_name:
                category = _hotspot_override(matched_name, category)

            results.append(LineResult(
                raw=raw, matched_name=matched_name, department=dept,
                category=category, amount=amount, score=score,
            ))

        return CategorizedBill(
            bill_name=bill_name, lines=results,
            categories=list(self.category_map.categories),
        )

    # ---- Render -----------------------------------------------------------

    def _render_bill(self) -> None:
        self.lines_tree.delete(*self.lines_tree.get_children())

        if self.current_bill is None:
            return

        # Update category cards
        totals = self.current_bill.totals()
        overhead = totals.pop("_overhead_split", 0.0)
        for cat, lbl in self._category_labels.items():
            val = totals.get(cat, 0.0)
            lbl.config(text=f"${val:,.2f}")

        # Grand total
        gt = self.current_bill.grand_total()
        if self._grand_total_label:
            self._grand_total_label.config(text=f"${gt:,.2f}")
        if overhead and self._overhead_label:
            n = len(self.current_bill.categories)
            self._overhead_label.config(
                text=f"Includes ${overhead:,.2f} overhead split across {n} categories"
            )
        else:
            self._overhead_label.config(text="")

        # Line items
        matched_count = 0
        for i, ln in enumerate(self.current_bill.lines):
            tag = "oddrow" if i % 2 else "evenrow"
            if not ln.matched_name:
                tag = "unmatched"
            else:
                matched_count += 1
            self.lines_tree.insert(
                "", "end",
                values=(
                    ln.matched_name or "(unmatched)",
                    ln.category,
                    f"${ln.amount:,.2f}" if ln.amount is not None else "—",
                ),
                tags=(tag,),
            )

        unmatched_count = len(self.current_bill.lines) - matched_count
        parts = [f"{len(self.current_bill.lines)} lines"]
        if unmatched_count:
            parts.append(f"{unmatched_count} unmatched")
        self._line_count_lbl.config(text="  |  ".join(parts))

    # ---- Export -----------------------------------------------------------

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
        self._set_status(f"Excel saved: {path}")
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
        self._set_status(f"PDF saved: {path}")
        messagebox.showinfo("Exported", f"Saved to:\n{path}")

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_var.set(text))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    if _DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

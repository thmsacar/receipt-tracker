#!/usr/bin/env python3
"""Generate Excel expense tracking workbook from receipts.json.

Creates a structured workbook (depenses.xlsx) with four sheets:
1. Synthèse Globale: KPIs, food vs non-food breakdown, store summary, and charts.
2. Articles Alimentaires: Dedicated food items table with conditional formatting.
3. Tickets de Caisse: Chronological receipt register.
4. Tous les Articles: Complete item-level expense log.
"""

import difflib
import json
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import xlsxwriter
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet

# Disable bytecode generation (__pycache__)
sys.dont_write_bytecode = True

# Taxonomy and normalizers from cli.py (with standalone fallback)
try:
    from cli import (
        CANONICAL_MARKET_MAP,
        CATEGORIES_ALIMENTAIRES,
        CATEGORIES_NON_ALIMENTAIRES,
        CATEGORIES_VALIDES as CATEGORIES_FR,
        normalize_category as normalize_category_fr,
        normalize_market_name,
    )
except ImportError:
    CATEGORIES_ALIMENTAIRES = [
        "Épicerie & Base",
        "Boucherie & Poissonnerie",
        "Fruits & Légumes",
        "Boissons & Snacks",
    ]
    CATEGORIES_NON_ALIMENTAIRES = [
        "Entretien & Maison",
        "Hygiène & Beauté",
        "Autre",
    ]
    CATEGORIES_FR = CATEGORIES_ALIMENTAIRES + CATEGORIES_NON_ALIMENTAIRES

    CANONICAL_MARKET_MAP = {
        "acem": "Boucherie Acem",
        "boucherie ace": "Boucherie Acem",
        "boucherie acem": "Boucherie Acem",
        "boucherie saint bruno": "Boucherie Saint Bruno",
        "boucherie st bruno": "Boucherie Saint Bruno",
        "saint bruno": "Boucherie Saint Bruno",
        "st bruno": "Boucherie Saint Bruno",
        "carrefour express": "Carrefour Express",
        "carrefour": "Carrefour",
        "lidl": "Lidl",
        "monoprix": "Monoprix",
        "auchan": "Auchan",
        "leclerc": "E.Leclerc",
        "e.leclerc": "E.Leclerc",
        "intermarche": "Intermarché",
        "intermarché": "Intermarché",
        "casino": "Casino",
        "franprix": "Franprix",
        "aldi": "Aldi",
        "action": "Action",
        "ikea": "Ikea",
        "h&m": "H&M",
        "decathlon": "Decathlon",
        "manav / primeur": "Manav / Primeur",
        "primeur": "Manav / Primeur",
        "manav": "Manav / Primeur",
    }

    def normalize_market_name(market: Optional[str], existing_markets: Optional[Iterable[str]] = None) -> str:
        """Normalize store name to canonical brand."""
        if not market or not str(market).strip():
            return "Commerce local"
        clean = re.sub(r"\(.*?\)", "", str(market)).strip()
        clean = re.sub(r"\s+", " ", clean)
        clean_lower = clean.lower()
        for key in sorted(CANONICAL_MARKET_MAP.keys(), key=len, reverse=True):
            if key in clean_lower:
                return CANONICAL_MARKET_MAP[key]
        if clean_lower.startswith("boucherie ace") or clean_lower.startswith("acem"):
            return "Boucherie Acem"
        all_targets = set(CANONICAL_MARKET_MAP.values())
        if existing_markets:
            for em in existing_markets:
                if em and str(em).strip():
                    all_targets.add(str(em).strip())
        matches = difflib.get_close_matches(clean, list(all_targets), n=1, cutoff=0.75)
        return matches[0] if matches else clean.title()

    def normalize_category_fr(cat: Optional[str]) -> str:
        """Normalize category string to the allowed taxonomy."""
        if not cat:
            return "Autre"
        cat_str = cat.strip()
        if cat_str in CATEGORIES_FR:
            return cat_str
        for v in CATEGORIES_FR:
            if v.lower() == cat_str.lower():
                return v
        return "Autre"


# ==============================================================================
# Excel Styles and Format Registry
# ==============================================================================

class ExcelFormatRegistry:
    """Centralized XlsxWriter format and color palette registry."""

    NAVY: str = "#1F4E79"
    ACCENT_BLUE: str = "#2E75B6"
    ZEBRA: str = "#F8FAFC"
    BORDER_COLOR: str = "#CBD5E1"
    FONT_FAMILY: str = "Segoe UI"

    FOOD_DARK: str = "#1B5E20"
    FOOD_HEADER: str = "#2E7D32"
    FOOD_ACCENT: str = "#388E3C"
    FOOD_BG_LIGHT: str = "#E8F5E9"
    FOOD_BORDER: str = "#A5D6A7"

    NONFOOD_HEADER: str = "#475569"
    NONFOOD_BG_LIGHT: str = "#F1F5F9"

    def __init__(self, workbook: Workbook) -> None:
        self.workbook = workbook
        self._build_formats()

    def _build_formats(self) -> None:
        """Create reusable workbook formats."""
        wb = self.workbook
        font = self.FONT_FAMILY

        # Titles and section headers
        self.title = wb.add_format({
            "font_name": font, "font_size": 16, "bold": True,
            "color": "#FFFFFF", "bg_color": self.NAVY,
            "align": "left", "valign": "vcenter",
        })
        self.subtitle = wb.add_format({
            "font_name": font, "font_size": 10, "italic": True,
            "color": "#E0E7FF", "bg_color": self.NAVY,
            "align": "left", "valign": "vcenter",
        })
        self.section_hdr = wb.add_format({
            "font_name": font, "font_size": 12, "bold": True,
            "color": self.NAVY, "bottom": 2, "bottom_color": self.ACCENT_BLUE,
            "valign": "vcenter",
        })
        self.section_hdr_food = wb.add_format({
            "font_name": font, "font_size": 12, "bold": True,
            "color": self.FOOD_DARK, "bottom": 2, "bottom_color": self.FOOD_HEADER,
            "valign": "vcenter",
        })

        # Table headers
        self.tbl_hdr = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": "#FFFFFF", "bg_color": self.NAVY,
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": self.NAVY, "text_wrap": True,
        })
        self.tbl_hdr_left = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": "#FFFFFF", "bg_color": self.NAVY,
            "align": "left", "valign": "vcenter",
            "border": 1, "border_color": self.NAVY, "text_wrap": True,
        })
        self.tbl_hdr_food = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": "#FFFFFF", "bg_color": self.FOOD_HEADER,
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": self.FOOD_HEADER, "text_wrap": True,
        })
        self.tbl_hdr_food_left = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": "#FFFFFF", "bg_color": self.FOOD_HEADER,
            "align": "left", "valign": "vcenter",
            "border": 1, "border_color": self.FOOD_HEADER, "text_wrap": True,
        })
        self.tbl_hdr_nonfood = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": "#FFFFFF", "bg_color": self.NONFOOD_HEADER,
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": self.NONFOOD_HEADER, "text_wrap": True,
        })
        self.tbl_hdr_nonfood_left = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": "#FFFFFF", "bg_color": self.NONFOOD_HEADER,
            "align": "left", "valign": "vcenter",
            "border": 1, "border_color": self.NONFOOD_HEADER, "text_wrap": True,
        })

        # Standard data cells
        self.cell_left = wb.add_format({
            "font_name": font, "font_size": 9, "align": "left", "valign": "vcenter",
            "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.cell_center = wb.add_format({
            "font_name": font, "font_size": 9, "align": "center", "valign": "vcenter",
            "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.qty = wb.add_format({
            "font_name": font, "font_size": 9, "align": "right", "valign": "vcenter",
            "num_format": "#,##0.###", "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.currency = wb.add_format({
            "font_name": font, "font_size": 9, "align": "right", "valign": "vcenter",
            "num_format": "€ #,##0.00", "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.percent = wb.add_format({
            "font_name": font, "font_size": 9, "align": "right", "valign": "vcenter",
            "num_format": "0.0%", "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })

        # Alternating (zebra) rows
        self.cell_left_zebra = wb.add_format({
            "font_name": font, "font_size": 9, "bg_color": self.ZEBRA,
            "align": "left", "valign": "vcenter",
            "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.cell_center_zebra = wb.add_format({
            "font_name": font, "font_size": 9, "bg_color": self.ZEBRA,
            "align": "center", "valign": "vcenter",
            "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.qty_zebra = wb.add_format({
            "font_name": font, "font_size": 9, "bg_color": self.ZEBRA,
            "align": "right", "valign": "vcenter",
            "num_format": "#,##0.###", "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.currency_zebra = wb.add_format({
            "font_name": font, "font_size": 9, "bg_color": self.ZEBRA,
            "align": "right", "valign": "vcenter",
            "num_format": "€ #,##0.00", "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })
        self.percent_zebra = wb.add_format({
            "font_name": font, "font_size": 9, "bg_color": self.ZEBRA,
            "align": "right", "valign": "vcenter",
            "num_format": "0.0%", "bottom": 1, "bottom_color": self.BORDER_COLOR,
        })

        # KPI cards - Row 1 (global metrics)
        self.kpi1_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": "#1B365D", "align": "center", "valign": "vcenter",
        })
        self.kpi1_curr = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": "#1B365D",
            "num_format": "€ #,##0.00", "bg_color": "#EDF4FA",
            "align": "center", "valign": "vcenter", "border": 1, "border_color": "#B8D5E5",
        })
        self.kpi2_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": "#2E75B6", "align": "center", "valign": "vcenter",
        })
        self.kpi2_val = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": "#2E75B6",
            "num_format": "#,##0", "bg_color": "#F0F6FB",
            "align": "center", "valign": "vcenter", "border": 1, "border_color": "#C2DCF0",
        })
        self.kpi3_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": "#4B6F96", "align": "center", "valign": "vcenter",
        })
        self.kpi3_val = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": "#4B6F96",
            "num_format": "#,##0", "bg_color": "#F4F7FA",
            "align": "center", "valign": "vcenter", "border": 1, "border_color": "#D3DFEC",
        })
        self.kpi4_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": "#708090", "align": "center", "valign": "vcenter",
        })
        self.kpi4_curr = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": "#708090",
            "num_format": "€ #,##0.00", "bg_color": "#F8FAFC",
            "align": "center", "valign": "vcenter", "border": 1, "border_color": "#CBD5E1",
        })

        # KPI cards - Row 2 (food vs non-food)
        self.kpi_food_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": self.FOOD_DARK, "align": "center", "valign": "vcenter",
        })
        self.kpi_food_curr = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": self.FOOD_DARK,
            "num_format": "€ #,##0.00", "bg_color": self.FOOD_BG_LIGHT,
            "align": "center", "valign": "vcenter", "border": 1, "border_color": self.FOOD_BORDER,
        })
        self.kpi_food_pct_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": self.FOOD_HEADER, "align": "center", "valign": "vcenter",
        })
        self.kpi_food_pct_val = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": self.FOOD_HEADER,
            "num_format": "0.0%", "bg_color": self.FOOD_BG_LIGHT,
            "align": "center", "valign": "vcenter", "border": 1, "border_color": self.FOOD_BORDER,
        })
        self.kpi_food_qty_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": self.FOOD_ACCENT, "align": "center", "valign": "vcenter",
        })
        self.kpi_food_qty_val = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": self.FOOD_ACCENT,
            "num_format": "#,##0", "bg_color": self.FOOD_BG_LIGHT,
            "align": "center", "valign": "vcenter", "border": 1, "border_color": self.FOOD_BORDER,
        })
        self.kpi_nonfood_title = wb.add_format({
            "font_name": font, "font_size": 9, "bold": True, "color": "#FFFFFF",
            "bg_color": self.NONFOOD_HEADER, "align": "center", "valign": "vcenter",
        })
        self.kpi_nonfood_curr = wb.add_format({
            "font_name": font, "font_size": 14, "bold": True, "color": self.NONFOOD_HEADER,
            "num_format": "€ #,##0.00", "bg_color": self.NONFOOD_BG_LIGHT,
            "align": "center", "valign": "vcenter", "border": 1, "border_color": "#CBD5E1",
        })

        # Grand totals
        self.total_lbl = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "bg_color": "#EAEEF3", "top": 1, "bottom": 2,
            "top_color": self.NAVY, "bottom_color": self.NAVY, "valign": "vcenter",
        })
        self.total_curr = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "num_format": "€ #,##0.00", "bg_color": "#EAEEF3",
            "top": 1, "bottom": 2, "top_color": self.NAVY, "bottom_color": self.NAVY,
            "align": "right", "valign": "vcenter",
        })
        self.total_qty = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "num_format": "#,##0.###", "bg_color": "#EAEEF3",
            "top": 1, "bottom": 2, "top_color": self.NAVY, "bottom_color": self.NAVY,
            "align": "right", "valign": "vcenter",
        })
        self.total_pct = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "num_format": "0.0%", "bg_color": "#EAEEF3",
            "top": 1, "bottom": 2, "top_color": self.NAVY, "bottom_color": self.NAVY,
            "align": "right", "valign": "vcenter",
        })

        # Food subtotals
        self.food_subtotal_lbl = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.FOOD_DARK, "bg_color": self.FOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.FOOD_HEADER, "bottom_color": self.FOOD_HEADER,
            "valign": "vcenter",
        })
        self.food_subtotal_curr = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.FOOD_DARK, "num_format": "€ #,##0.00", "bg_color": self.FOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.FOOD_HEADER, "bottom_color": self.FOOD_HEADER,
            "align": "right", "valign": "vcenter",
        })
        self.food_subtotal_qty = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.FOOD_DARK, "num_format": "#,##0.###", "bg_color": self.FOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.FOOD_HEADER, "bottom_color": self.FOOD_HEADER,
            "align": "right", "valign": "vcenter",
        })
        self.food_subtotal_pct = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.FOOD_DARK, "num_format": "0.0%", "bg_color": self.FOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.FOOD_HEADER, "bottom_color": self.FOOD_HEADER,
            "align": "right", "valign": "vcenter",
        })

        # Non-food subtotals
        self.nonfood_subtotal_lbl = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.NONFOOD_HEADER, "bg_color": self.NONFOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.NONFOOD_HEADER, "bottom_color": self.NONFOOD_HEADER,
            "valign": "vcenter",
        })
        self.nonfood_subtotal_curr = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.NONFOOD_HEADER, "num_format": "€ #,##0.00", "bg_color": self.NONFOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.NONFOOD_HEADER, "bottom_color": self.NONFOOD_HEADER,
            "align": "right", "valign": "vcenter",
        })
        self.nonfood_subtotal_qty = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.NONFOOD_HEADER, "num_format": "#,##0.###", "bg_color": self.NONFOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.NONFOOD_HEADER, "bottom_color": self.NONFOOD_HEADER,
            "align": "right", "valign": "vcenter",
        })
        self.nonfood_subtotal_pct = wb.add_format({
            "font_name": font, "font_size": 10, "bold": True,
            "color": self.NONFOOD_HEADER, "num_format": "0.0%", "bg_color": self.NONFOOD_BG_LIGHT,
            "top": 1, "bottom": 2, "top_color": self.NONFOOD_HEADER, "bottom_color": self.NONFOOD_HEADER,
            "align": "right", "valign": "vcenter",
        })


# ==============================================================================
# Financial Data Aggregation
# ==============================================================================

@dataclass
class ExpenseDataSummary:
    """Aggregated expense data container."""

    sorted_receipts: List[Dict[str, Any]]
    all_items_rows: List[Dict[str, Any]]
    food_items_rows: List[Dict[str, Any]]
    cat_stats: Dict[str, Dict[str, Any]]
    market_stats: Dict[str, Dict[str, Any]]

    total_spend_all: float
    total_items_all: int
    total_receipts_all: int
    avg_receipt_all: float

    food_spend_all: float
    food_items_all: int
    food_pct_all: float

    nonfood_spend_all: float
    nonfood_items_all: int
    nonfood_pct_all: float


class ExpenseAggregator:
    """Aggregate expense data across categories and stores."""

    @staticmethod
    def process(raw_receipts: List[Dict[str, Any]]) -> ExpenseDataSummary:
        """Parse, sort, and aggregate receipt expenses."""
        def sort_key(r: Dict[str, Any]) -> str:
            d = str(r.get("date") or "9999-99-99")
            if len(d) == 7:
                d = f"{d}-01"
            return d

        data = sorted(raw_receipts, key=sort_key)

        total_spend_all = 0.0
        total_items_all = 0
        total_receipts_all = len(data)

        food_spend_all = 0.0
        food_items_all = 0

        nonfood_spend_all = 0.0
        nonfood_items_all = 0

        cat_stats: Dict[str, Dict[str, Any]] = {c: {"count": 0, "total": 0.0} for c in CATEGORIES_FR}
        market_stats: Dict[str, Dict[str, Any]] = defaultdict(lambda: {
            "receipts": 0,
            "total": 0.0,
            "items": 0,
            "branch": "",
        })

        all_items_rows: List[Dict[str, Any]] = []
        food_items_rows: List[Dict[str, Any]] = []

        for idx, r in enumerate(data, start=1):
            date_str = str(r.get("date") or "-")
            r_id = r.get("id") or f"{date_str}_{idx:02d}"
            m = normalize_market_name(r.get("market", ""))
            branch = r.get("branch") or ""
            tot = float(r.get("total_amount") or 0.0)

            items_list = r.get("items") or []
            items_cnt = len(items_list)

            total_spend_all += tot
            total_items_all += items_cnt

            market_stats[m]["receipts"] += 1
            market_stats[m]["total"] += tot
            market_stats[m]["items"] += items_cnt
            if branch and not market_stats[m]["branch"]:
                market_stats[m]["branch"] = branch

            for u in items_list:
                c = normalize_category_fr(u.get("category"))
                price = float(u.get("total_price") or 0.0)
                qty = float(u.get("quantity", 1.0))
                unit = u.get("unit") or "pièce"
                if unit == "adet":
                    unit = "pièce"
                u_price = float(u.get("unit_price") or 0.0)
                desc = u.get("description") or u.get("raw_text") or ""
                raw = u.get("raw_text") or ""

                cat_stats[c]["count"] += 1
                cat_stats[c]["total"] += price

                item_dict = {
                    "r_id": r_id,
                    "market": m,
                    "branch": branch,
                    "date": date_str,
                    "desc": desc,
                    "raw": raw,
                    "category": c,
                    "qty": qty,
                    "unit": unit,
                    "u_price": u_price,
                    "t_price": price,
                    "currency": r.get("currency", "EUR"),
                }
                all_items_rows.append(item_dict)

                if c in CATEGORIES_ALIMENTAIRES:
                    food_spend_all += price
                    food_items_all += 1
                    food_items_rows.append(item_dict)
                else:
                    nonfood_spend_all += price
                    nonfood_items_all += 1

        avg_receipt_all = (total_spend_all / total_receipts_all) if total_receipts_all else 0.0
        food_pct_all = (food_spend_all / total_spend_all) if total_spend_all else 0.0
        nonfood_pct_all = (nonfood_spend_all / total_spend_all) if total_spend_all else 0.0

        return ExpenseDataSummary(
            sorted_receipts=data,
            all_items_rows=all_items_rows,
            food_items_rows=food_items_rows,
            cat_stats=cat_stats,
            market_stats=market_stats,
            total_spend_all=total_spend_all,
            total_items_all=total_items_all,
            total_receipts_all=total_receipts_all,
            avg_receipt_all=avg_receipt_all,
            food_spend_all=food_spend_all,
            food_items_all=food_items_all,
            food_pct_all=food_pct_all,
            nonfood_spend_all=nonfood_spend_all,
            nonfood_items_all=nonfood_items_all,
            nonfood_pct_all=nonfood_pct_all,
        )


# ==============================================================================
# Worksheet Writers
# ==============================================================================

def _write_summary_worksheet(
    workbook: Workbook,
    ws: Worksheet,
    summary: ExpenseDataSummary,
    fmt: ExcelFormatRegistry,
) -> None:
    """Build the global summary dashboard worksheet with KPIs, tables, and charts."""
    ws.hide_gridlines(0)

    # Title and subtitle banner
    ws.merge_range("A1:L1", "  RAPPORT DE SUIVI DES DÉPENSES & COURSES (FRANCE)", fmt.title)
    ws.merge_range("A2:L2", "  Analyse globale avec distinction nette entre Dépenses Alimentaires (Comestibles) et Non-Alimentaires", fmt.subtitle)
    ws.set_row(0, 24)
    ws.set_row(1, 18)

    # KPI cards - Row 1: Global metrics
    ws.set_row(3, 16)
    ws.set_row(4, 26)

    ws.merge_range("A4:B4", "TOTAL DÉPENSÉ (GLOBAL)", fmt.kpi1_title)
    ws.write_formula("A5", f"='🧾 Tickets de Caisse'!H{summary.total_receipts_all + 2}", fmt.kpi1_curr, value=round(summary.total_spend_all, 2))
    ws.write_blank("B5", None, fmt.kpi1_curr)

    ws.merge_range("D4:E4", "NOMBRE DE TICKETS", fmt.kpi2_title)
    ws.write_formula("D5", f"=COUNTA('🧾 Tickets de Caisse'!A2:A{summary.total_receipts_all + 1})", fmt.kpi2_val, value=summary.total_receipts_all)
    ws.write_blank("E5", None, fmt.kpi2_val)

    ws.merge_range("G4:H4", "ARTICLES ACHETÉS (TOTAL)", fmt.kpi3_title)
    ws.write_formula("G5", f"=COUNTA('🛒 Tous les Articles'!A2:A{summary.total_items_all + 1})", fmt.kpi3_val, value=summary.total_items_all)
    ws.write_blank("H5", None, fmt.kpi3_val)

    ws.merge_range("J4:K4", "PANIER MOYEN / TICKET", fmt.kpi4_title)
    ws.write_formula("J5", "=A5/D5", fmt.kpi4_curr, value=round(summary.avg_receipt_all, 2))
    ws.write_blank("K5", None, fmt.kpi4_curr)

    # KPI cards - Row 2: Food vs non-food focus
    ws.set_row(6, 16)
    ws.set_row(7, 26)

    ws.merge_range("A7:B7", "🥗 BUDGET ALIMENTAIRE (COMESTIBLE)", fmt.kpi_food_title)
    ws.write_formula("A8", "=C16", fmt.kpi_food_curr, value=round(summary.food_spend_all, 2))
    ws.write_blank("B8", None, fmt.kpi_food_curr)

    ws.merge_range("D7:E7", "PART ALIMENTAIRE / TOTAL", fmt.kpi_food_pct_title)
    ws.write_formula("D8", "=A8/A5", fmt.kpi_food_pct_val, value=summary.food_pct_all)
    ws.write_blank("E8", None, fmt.kpi_food_pct_val)

    ws.merge_range("G7:H7", "ARTICLES COMESTIBLES", fmt.kpi_food_qty_title)
    ws.write_formula("G8", "=B16", fmt.kpi_food_qty_val, value=summary.food_items_all)
    ws.write_blank("H8", None, fmt.kpi_food_qty_val)

    ws.merge_range("J7:K7", "🛍️ BUDGET NON-ALIMENTAIRE", fmt.kpi_nonfood_title)
    ws.write_formula("J8", "=C23", fmt.kpi_nonfood_curr, value=round(summary.nonfood_spend_all, 2))
    ws.write_blank("K8", None, fmt.kpi_nonfood_curr)

    # Table 1: Food spending breakdown
    row = 9
    ws.write(row, 0, "🥗 1. DÉPENSES ALIMENTAIRES (COMESTIBLES : ÉPICERIE, BOUCHERIE, FRUITS & LÉGUMES, BOISSONS)", fmt.section_hdr_food)
    row += 1
    food_headers = ["Catégorie Alimentaire", "Articles", "Total Dépensé (€)", "Part Alimentaire (%)", "Part Budget Total (%)", "Prix Moyen / Art. (€)"]
    for c_idx, h in enumerate(food_headers):
        ws.write(row, c_idx, h, fmt.tbl_hdr_food_left if c_idx == 0 else fmt.tbl_hdr_food)
    ws.set_row(row, 22)

    sorted_food_cats = sorted(
        [(c, summary.cat_stats[c]) for c in CATEGORIES_ALIMENTAIRES],
        key=lambda x: x[1]["total"],
        reverse=True,
    )

    for i, (cat_name, st) in enumerate(sorted_food_cats):
        row += 1
        is_z = (i % 2 == 1)
        c_left = fmt.cell_left_zebra if is_z else fmt.cell_left
        c_qty = fmt.qty_zebra if is_z else fmt.qty
        c_curr = fmt.currency_zebra if is_z else fmt.currency
        c_pct = fmt.percent_zebra if is_z else fmt.percent

        ws.write(row, 0, cat_name, c_left)
        ws.write(row, 1, st["count"], c_qty)
        ws.write(row, 2, st["total"], c_curr)

        pct_in_food = (st["total"] / summary.food_spend_all) if summary.food_spend_all else 0.0
        ws.write_formula(row, 3, f"=C{row+1}/$C$16", c_pct, value=pct_in_food)

        pct_in_total = (st["total"] / summary.total_spend_all) if summary.total_spend_all else 0.0
        ws.write_formula(row, 4, f"=C{row+1}/$A$5", c_pct, value=pct_in_total)

        avg_art = (st["total"] / st["count"]) if st["count"] else 0.0
        ws.write_formula(row, 5, f"=C{row+1}/B{row+1}", c_curr, value=round(avg_art, 2))
        ws.set_row(row, 18)

    # Food subtotal row
    row += 1
    ws.write(row, 0, "Sous-Total Alimentaire (Comestibles)", fmt.food_subtotal_lbl)
    ws.write_formula(row, 1, "=SUM(B12:B15)", fmt.food_subtotal_qty, value=summary.food_items_all)
    ws.write_formula(row, 2, "=SUM(C12:C15)", fmt.food_subtotal_curr, value=round(summary.food_spend_all, 2))
    ws.write_formula(row, 3, "=SUM(D12:D15)", fmt.food_subtotal_pct, value=1.0)
    ws.write_formula(row, 4, "=C16/$A$5", fmt.food_subtotal_pct, value=summary.food_pct_all)
    avg_food_item = (summary.food_spend_all / summary.food_items_all) if summary.food_items_all else 0.0
    ws.write_formula(row, 5, "=C16/B16", fmt.food_subtotal_curr, value=round(avg_food_item, 2))
    ws.set_row(row, 20)

    # Table 2: Non-Food spending breakdown
    row += 2
    ws.write(row, 0, "🛍️ 2. DÉPENSES NON-ALIMENTAIRES (MAISON, HYGIÈNE, BRICOLAGE & DIVERS)", fmt.section_hdr)
    row += 1
    nonfood_headers = ["Catégorie Non-Alimentaire", "Articles", "Total Dépensé (€)", "Part Non-Alim. (%)", "Part Budget Total (%)", "Prix Moyen / Art. (€)"]
    for c_idx, h in enumerate(nonfood_headers):
        ws.write(row, c_idx, h, fmt.tbl_hdr_nonfood_left if c_idx == 0 else fmt.tbl_hdr_nonfood)
    ws.set_row(row, 22)

    sorted_nonfood_cats = sorted(
        [(c, summary.cat_stats[c]) for c in CATEGORIES_NON_ALIMENTAIRES],
        key=lambda x: x[1]["total"],
        reverse=True,
    )

    for i, (cat_name, st) in enumerate(sorted_nonfood_cats):
        row += 1
        is_z = (i % 2 == 1)
        c_left = fmt.cell_left_zebra if is_z else fmt.cell_left
        c_qty = fmt.qty_zebra if is_z else fmt.qty
        c_curr = fmt.currency_zebra if is_z else fmt.currency
        c_pct = fmt.percent_zebra if is_z else fmt.percent

        ws.write(row, 0, cat_name, c_left)
        ws.write(row, 1, st["count"], c_qty)
        ws.write(row, 2, st["total"], c_curr)

        pct_in_nonfood = (st["total"] / summary.nonfood_spend_all) if summary.nonfood_spend_all else 0.0
        ws.write_formula(row, 3, f"=C{row+1}/$C$23", c_pct, value=pct_in_nonfood)

        pct_in_total = (st["total"] / summary.total_spend_all) if summary.total_spend_all else 0.0
        ws.write_formula(row, 4, f"=C{row+1}/$A$5", c_pct, value=pct_in_total)

        avg_art = (st["total"] / st["count"]) if st["count"] else 0.0
        ws.write_formula(row, 5, f"=C{row+1}/B{row+1}", c_curr, value=round(avg_art, 2))
        ws.set_row(row, 18)

    # Non-food subtotal row
    row += 1
    ws.write(row, 0, "Sous-Total Non-Alimentaire", fmt.nonfood_subtotal_lbl)
    ws.write_formula(row, 1, "=SUM(B20:B22)", fmt.nonfood_subtotal_qty, value=summary.nonfood_items_all)
    ws.write_formula(row, 2, "=SUM(C20:C22)", fmt.nonfood_subtotal_curr, value=round(summary.nonfood_spend_all, 2))
    ws.write_formula(row, 3, "=SUM(D20:D22)", fmt.nonfood_subtotal_pct, value=1.0)
    ws.write_formula(row, 4, "=C23/$A$5", fmt.nonfood_subtotal_pct, value=summary.nonfood_pct_all)
    avg_nonfood_item = (summary.nonfood_spend_all / summary.nonfood_items_all) if summary.nonfood_items_all else 0.0
    ws.write_formula(row, 5, "=C23/B23", fmt.nonfood_subtotal_curr, value=round(avg_nonfood_item, 2))
    ws.set_row(row, 20)

    # Table 3: Macro summary (Food vs Non-Food)
    row += 2
    ws.write(row, 0, "⚖️ 3. SYNTHÈSE MACRO : ALIMENTAIRE (COMESTIBLE) VS NON-ALIMENTAIRE", fmt.section_hdr)
    row += 1
    macro_headers = ["Secteur de Dépense", "Articles", "Total Dépensé (€)", "Part du Budget (%)", "Prix Moyen / Art. (€)"]
    for c_idx, h in enumerate(macro_headers):
        ws.write(row, c_idx, h, fmt.tbl_hdr_left if c_idx == 0 else fmt.tbl_hdr)
    ws.set_row(row, 22)

    # Food line
    row += 1
    ws.write(row, 0, "🥗 Alimentaire & Comestibles (Épicerie, Boucherie, Fruits, Boissons)", fmt.cell_left)
    ws.write_formula(row, 1, "=B16", fmt.qty, value=summary.food_items_all)
    ws.write_formula(row, 2, "=C16", fmt.currency, value=round(summary.food_spend_all, 2))
    ws.write_formula(row, 3, "=C27/$C$29", fmt.percent, value=summary.food_pct_all)
    ws.write_formula(row, 4, "=C27/B27", fmt.currency, value=round(avg_food_item, 2))
    ws.set_row(row, 18)

    # Non-food line
    row += 1
    ws.write(row, 0, "🛍️ Non-Alimentaire (Entretien, Maison, Hygiène, Divers)", fmt.cell_left_zebra)
    ws.write_formula(row, 1, "=B23", fmt.qty_zebra, value=summary.nonfood_items_all)
    ws.write_formula(row, 2, "=C23", fmt.currency_zebra, value=round(summary.nonfood_spend_all, 2))
    ws.write_formula(row, 3, "=C28/$C$29", fmt.percent_zebra, value=summary.nonfood_pct_all)
    ws.write_formula(row, 4, "=C28/B28", fmt.currency_zebra, value=round(avg_nonfood_item, 2))
    ws.set_row(row, 18)

    # Grand total row
    row += 1
    ws.write(row, 0, "TOTAL GÉNÉRAL", fmt.total_lbl)
    ws.write_formula(row, 1, "=SUM(B27:B28)", fmt.total_qty, value=summary.total_items_all)
    ws.write_formula(row, 2, "=SUM(C27:C28)", fmt.total_curr, value=round(summary.total_spend_all, 2))
    ws.write_formula(row, 3, "=SUM(D27:D28)", fmt.total_pct, value=1.0)
    overall_art_avg = (summary.total_spend_all / summary.total_items_all) if summary.total_items_all else 0.0
    ws.write_formula(row, 4, "=C29/B29", fmt.total_curr, value=round(overall_art_avg, 2))
    ws.set_row(row, 20)

    # Table 4: Store breakdown
    row += 2
    ws.write(row, 0, "🏪 4. RÉPARTITION PAR ENSEIGNE / MAGASIN", fmt.section_hdr)
    row += 1
    mkt_headers = ["Magasin", "Localisation Principale", "Tickets", "Articles", "Montant Total (€)", "Part (%)"]
    for c_idx, h in enumerate(mkt_headers):
        ws.write(row, c_idx, h, fmt.tbl_hdr_left if c_idx in [0, 1] else fmt.tbl_hdr)
    ws.set_row(row, 22)

    mkt_start_row = row + 2
    sorted_markets = sorted(summary.market_stats.items(), key=lambda x: x[1]["total"], reverse=True)
    for i, (m, st) in enumerate(sorted_markets):
        row += 1
        is_z = (i % 2 == 1)
        c_left = fmt.cell_left_zebra if is_z else fmt.cell_left
        c_center = fmt.cell_center_zebra if is_z else fmt.cell_center
        c_curr = fmt.currency_zebra if is_z else fmt.currency
        c_pct = fmt.percent_zebra if is_z else fmt.percent
        c_qty = fmt.qty_zebra if is_z else fmt.qty

        ws.write(row, 0, m, c_left)
        ws.write(row, 1, st["branch"], c_left)
        ws.write(row, 2, st["receipts"], c_center)
        ws.write(row, 3, st["items"], c_qty)
        ws.write(row, 4, st["total"], c_curr)
        total_cell_ref = f"$E${mkt_start_row + len(sorted_markets)}"
        pct_val = st["total"] / summary.total_spend_all if summary.total_spend_all else 0.0
        ws.write_formula(row, 5, f"=E{row+1}/{total_cell_ref}", c_pct, value=pct_val)
        ws.set_row(row, 18)

    row += 1
    mkt_end_row = row
    ws.write(row, 0, "Total Général Enseignes", fmt.total_lbl)
    ws.write(row, 1, "", fmt.total_lbl)
    ws.write_formula(row, 2, f"=SUM(C{mkt_start_row}:C{mkt_end_row})", fmt.total_qty, value=summary.total_receipts_all)
    ws.write_formula(row, 3, f"=SUM(D{mkt_start_row}:D{mkt_end_row})", fmt.total_qty, value=summary.total_items_all)
    ws.write_formula(row, 4, f"=SUM(E{mkt_start_row}:E{mkt_end_row})", fmt.total_curr, value=round(summary.total_spend_all, 2))
    ws.write_formula(row, 5, f"=SUM(F{mkt_start_row}:F{mkt_end_row})", fmt.total_pct, value=1.0)
    ws.set_row(row, 20)

    # Column dimensions
    ws.set_column("A:A", 32)
    ws.set_column("B:B", 24)
    ws.set_column("C:C", 18)
    ws.set_column("D:D", 20)
    ws.set_column("E:E", 20)
    ws.set_column("F:F", 20)
    ws.set_column("G:G", 4)
    ws.set_column("H:H", 16)
    ws.set_column("I:I", 16)
    ws.set_column("J:J", 16)
    ws.set_column("K:K", 16)
    ws.set_column("L:L", 16)

    # Embedded native charts
    # Chart 1: Macro distribution (doughnut)
    chart_macro = workbook.add_chart({"type": "doughnut"})
    chart_macro.add_series({
        "name": "Macro-Répartition",
        "categories": ["📊 Synthèse Globale", 26, 0, 27, 0],
        "values": ["📊 Synthèse Globale", 26, 2, 27, 2],
        "points": [
            {"fill": {"color": "#2E7D32"}},
            {"fill": {"color": "#475569"}},
        ],
        "data_labels": {"percentage": True, "value": False},
    })
    chart_macro.set_title({"name": "Alimentaire (Comestible) vs Non-Alimentaire"})
    chart_macro.set_size({"width": 480, "height": 240})
    ws.insert_chart("H10", chart_macro)

    # Chart 2: Food categories (doughnut)
    chart_food = workbook.add_chart({"type": "doughnut"})
    chart_food.add_series({
        "name": "Détail Alimentaire",
        "categories": ["📊 Synthèse Globale", 11, 0, 14, 0],
        "values": ["📊 Synthèse Globale", 11, 2, 14, 2],
        "data_labels": {"percentage": True, "value": False},
    })
    chart_food.set_title({"name": "Répartition Dépenses Alimentaires (€)"})
    chart_food.set_size({"width": 480, "height": 240})
    ws.insert_chart("H23", chart_food)

    # Chart 3: Store spending (columns)
    chart_mkt = workbook.add_chart({"type": "column"})
    chart_mkt.add_series({
        "name": "Dépenses (€)",
        "categories": ["📊 Synthèse Globale", mkt_start_row - 1, 0, mkt_end_row - 1, 0],
        "values": ["📊 Synthèse Globale", mkt_start_row - 1, 4, mkt_end_row - 1, 4],
        "fill": {"color": "#1F4E79"},
        "data_labels": {"value": True, "num_format": "€ 0"},
    })
    chart_mkt.set_title({"name": "Dépenses par Enseigne (€)"})
    chart_mkt.set_legend({"none": True})
    chart_mkt.set_y_axis({"name": "Montant (€)", "major_gridlines": {"visible": True}})
    chart_mkt.set_x_axis({"name": "Enseigne"})
    chart_mkt.set_size({"width": 480, "height": 270})
    ws.insert_chart("H36", chart_mkt)


def _write_food_items_worksheet(
    ws: Worksheet,
    summary: ExpenseDataSummary,
    fmt: ExcelFormatRegistry,
) -> None:
    """Build dedicated food items worksheet with filtering."""
    ws.hide_gridlines(0)
    ws.freeze_panes(1, 0)

    food_headers = [
        "#",
        "ID Ticket",
        "Enseigne",
        "Localisation",
        "Date",
        "Description du Produit",
        "Texte Brut (Ticket)",
        "Catégorie Alimentaire",
        "Quantité",
        "Unité",
        "Prix Unitaire (€)",
        "Prix Total (€)",
        "Devise",
    ]

    for c_idx, h in enumerate(food_headers):
        is_ctr = c_idx in [0, 1, 4, 8, 9, 10, 11, 12]
        ws.write(0, c_idx, h, fmt.tbl_hdr_food if is_ctr else fmt.tbl_hdr_food_left)
    ws.set_row(0, 24)

    sum_food_qty = 0.0
    sum_food_price = 0.0

    for idx, it in enumerate(summary.food_items_rows, start=1):
        row = idx
        is_z = (idx % 2 == 0)

        c_left = fmt.cell_left_zebra if is_z else fmt.cell_left
        c_center = fmt.cell_center_zebra if is_z else fmt.cell_center
        c_curr = fmt.currency_zebra if is_z else fmt.currency
        c_qty = fmt.qty_zebra if is_z else fmt.qty

        sum_food_qty += it["qty"]
        sum_food_price += it["t_price"]

        ws.write(row, 0, idx, c_center)
        ws.write(row, 1, it["r_id"], c_center)
        ws.write(row, 2, it["market"], c_left)
        ws.write(row, 3, it["branch"], c_left)
        ws.write(row, 4, it["date"], c_center)
        ws.write(row, 5, it["desc"], c_left)
        ws.write(row, 6, it["raw"], c_left)
        ws.write(row, 7, it["category"], c_left)
        ws.write(row, 8, it["qty"], c_qty)
        ws.write(row, 9, it["unit"], c_center)
        ws.write(row, 10, it["u_price"], c_curr)
        ws.write(row, 11, it["t_price"], c_curr)
        ws.write(row, 12, it["currency"], c_center)
        ws.set_row(row, 19)

    tot_row = len(summary.food_items_rows) + 1
    ws.write(tot_row, 0, "TOTAL ALIMENTAIRE", fmt.food_subtotal_lbl)
    for c in range(1, 8):
        ws.write(tot_row, c, "", fmt.food_subtotal_lbl)
    ws.write_formula(tot_row, 8, f"=SUM(I2:I{tot_row})", fmt.food_subtotal_qty, value=round(sum_food_qty, 3))
    ws.write(tot_row, 9, "", fmt.food_subtotal_lbl)
    ws.write(tot_row, 10, "", fmt.food_subtotal_lbl)
    ws.write_formula(tot_row, 11, f"=SUM(L2:L{tot_row})", fmt.food_subtotal_curr, value=round(sum_food_price, 2))
    ws.write(tot_row, 12, "EUR", fmt.food_subtotal_lbl)
    ws.set_row(tot_row, 22)

    ws.set_column("A:A", 6)
    ws.set_column("B:B", 16)
    ws.set_column("C:C", 22)
    ws.set_column("D:D", 20)
    ws.set_column("E:E", 12)
    ws.set_column("F:F", 38)
    ws.set_column("G:G", 26)
    ws.set_column("H:H", 26)
    ws.set_column("I:I", 10)
    ws.set_column("J:J", 10)
    ws.set_column("K:K", 14)
    ws.set_column("L:L", 16)
    ws.set_column("M:M", 8)
    ws.autofilter(0, 0, len(summary.food_items_rows), len(food_headers) - 1)


def _write_receipts_worksheet(
    ws: Worksheet,
    summary: ExpenseDataSummary,
    fmt: ExcelFormatRegistry,
) -> None:
    """Build chronological receipt register worksheet."""
    ws.hide_gridlines(0)
    ws.freeze_panes(1, 0)

    rcpt_headers = [
        "#",
        "ID Chronologique",
        "Enseigne",
        "Localisation",
        "Date",
        "N° Reçu",
        "Articles",
        "Montant Total (€)",
        "Devise",
    ]

    for c_idx, h in enumerate(rcpt_headers):
        is_ctr = c_idx in [0, 1, 4, 5, 6, 7, 8]
        ws.write(0, c_idx, h, fmt.tbl_hdr if is_ctr else fmt.tbl_hdr_left)
    ws.set_row(0, 24)

    for i, r in enumerate(summary.sorted_receipts):
        row = i + 1
        is_z = (i % 2 == 1)
        c_left = fmt.cell_left_zebra if is_z else fmt.cell_left
        c_center = fmt.cell_center_zebra if is_z else fmt.cell_center
        c_curr = fmt.currency_zebra if is_z else fmt.currency
        c_qty = fmt.qty_zebra if is_z else fmt.qty

        r_id = r.get("id") or f"{r.get('date', '2026')}_{i+1:02d}"
        items_list = r.get("items") or []
        items_in_rcpt = len(items_list)
        tot_rcpt = float(r.get("total_amount") or 0.0)

        ws.write(row, 0, i + 1, c_center)
        ws.write(row, 1, r_id, c_center)
        ws.write(row, 2, normalize_market_name(r.get("market", "")), c_left)
        ws.write(row, 3, r.get("branch") or "", c_left)
        ws.write(row, 4, r.get("date") or "", c_center)
        ws.write(row, 5, r.get("receipt_no") or "-", c_center)
        ws.write(row, 6, items_in_rcpt, c_qty)
        ws.write(row, 7, tot_rcpt, c_curr)
        ws.write(row, 8, r.get("currency") or "EUR", c_center)
        ws.set_row(row, 20)

    tot_row = len(summary.sorted_receipts) + 1
    ws.write(tot_row, 0, "TOTAL", fmt.total_lbl)
    for c in range(1, 6):
        ws.write(tot_row, c, "", fmt.total_lbl)
    ws.write_formula(tot_row, 6, f"=SUM(G2:G{tot_row})", fmt.total_qty, value=summary.total_items_all)
    ws.write_formula(tot_row, 7, f"=SUM(H2:H{tot_row})", fmt.total_curr, value=round(summary.total_spend_all, 2))
    ws.write(tot_row, 8, "EUR", fmt.total_lbl)
    ws.set_row(tot_row, 22)

    ws.set_column("A:A", 6)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 30)
    ws.set_column("D:D", 24)
    ws.set_column("E:E", 14)
    ws.set_column("F:F", 14)
    ws.set_column("G:G", 12)
    ws.set_column("H:H", 18)
    ws.set_column("I:I", 10)
    ws.autofilter(0, 0, len(summary.sorted_receipts), len(rcpt_headers) - 1)


def _write_all_items_worksheet(
    ws: Worksheet,
    summary: ExpenseDataSummary,
    fmt: ExcelFormatRegistry,
) -> None:
    """Build complete item-level ledger worksheet."""
    ws.hide_gridlines(0)
    ws.freeze_panes(1, 0)

    all_headers = [
        "#",
        "ID Ticket",
        "Enseigne",
        "Localisation",
        "Date",
        "Description Normalisée",
        "Texte Brut (Ticket)",
        "Catégorie",
        "Quantité",
        "Unité",
        "Prix Unitaire (€)",
        "Prix Total (€)",
        "Devise",
    ]

    for c_idx, h in enumerate(all_headers):
        is_ctr = c_idx in [0, 1, 4, 8, 9, 10, 11, 12]
        ws.write(0, c_idx, h, fmt.tbl_hdr if is_ctr else fmt.tbl_hdr_left)
    ws.set_row(0, 24)

    sum_all_qty = 0.0
    sum_all_price = 0.0

    for idx, it in enumerate(summary.all_items_rows, start=1):
        row = idx
        is_z = (idx % 2 == 0)

        c_left = fmt.cell_left_zebra if is_z else fmt.cell_left
        c_center = fmt.cell_center_zebra if is_z else fmt.cell_center
        c_curr = fmt.currency_zebra if is_z else fmt.currency
        c_qty = fmt.qty_zebra if is_z else fmt.qty

        sum_all_qty += it["qty"]
        sum_all_price += it["t_price"]

        ws.write(row, 0, idx, c_center)
        ws.write(row, 1, it["r_id"], c_center)
        ws.write(row, 2, it["market"], c_left)
        ws.write(row, 3, it["branch"], c_left)
        ws.write(row, 4, it["date"], c_center)
        ws.write(row, 5, it["desc"], c_left)
        ws.write(row, 6, it["raw"], c_left)
        ws.write(row, 7, it["category"], c_left)
        ws.write(row, 8, it["qty"], c_qty)
        ws.write(row, 9, it["unit"], c_center)
        ws.write(row, 10, it["u_price"], c_curr)
        ws.write(row, 11, it["t_price"], c_curr)
        ws.write(row, 12, it["currency"], c_center)
        ws.set_row(row, 19)

    item_tot_row = len(summary.all_items_rows) + 1
    ws.write(item_tot_row, 0, "TOTAL GLOBAL", fmt.total_lbl)
    for c in range(1, 8):
        ws.write(item_tot_row, c, "", fmt.total_lbl)
    ws.write_formula(item_tot_row, 8, f"=SUM(I2:I{item_tot_row})", fmt.total_qty, value=round(sum_all_qty, 3))
    ws.write(item_tot_row, 9, "", fmt.total_lbl)
    ws.write(item_tot_row, 10, "", fmt.total_lbl)
    ws.write_formula(item_tot_row, 11, f"=SUM(L2:L{item_tot_row})", fmt.total_curr, value=round(sum_all_price, 2))
    ws.write(item_tot_row, 12, "EUR", fmt.total_lbl)
    ws.set_row(item_tot_row, 22)

    ws.set_column("A:A", 6)
    ws.set_column("B:B", 16)
    ws.set_column("C:C", 22)
    ws.set_column("D:D", 20)
    ws.set_column("E:E", 12)
    ws.set_column("F:F", 38)
    ws.set_column("G:G", 26)
    ws.set_column("H:H", 26)
    ws.set_column("I:I", 10)
    ws.set_column("J:J", 10)
    ws.set_column("K:K", 14)
    ws.set_column("L:L", 16)
    ws.set_column("M:M", 8)
    ws.autofilter(0, 0, len(summary.all_items_rows), len(all_headers) - 1)


# ==============================================================================
# Workbook Generation Orchestration
# ==============================================================================

def generate_excel(
    json_path: str = "receipts.json",
    output_path: str = "depenses.xlsx",
) -> None:
    """Generate depenses.xlsx from receipts.json."""
    with open(json_path, "r", encoding="utf-8") as f:
        raw_data = json.load(f)

    # 1. Aggregate and compute metrics
    summary = ExpenseAggregator.process(raw_data)

    # 2. Initialize workbook and formats
    workbook = xlsxwriter.Workbook(output_path, {"strings_to_numbers": True})
    formats = ExcelFormatRegistry(workbook)

    # 3. Build worksheets sequentially
    ws_summary = workbook.add_worksheet("📊 Synthèse Globale")
    _write_summary_worksheet(workbook, ws_summary, summary, formats)

    ws_food = workbook.add_worksheet("🥗 Articles Alimentaires")
    _write_food_items_worksheet(ws_food, summary, formats)

    ws_receipts = workbook.add_worksheet("🧾 Tickets de Caisse")
    _write_receipts_worksheet(ws_receipts, summary, formats)

    ws_items = workbook.add_worksheet("🛒 Tous les Articles")
    _write_all_items_worksheet(ws_items, summary, formats)

    # 4. Finalize and save
    workbook.close()
    print(f"Excel généré avec succès : {output_path}")


if __name__ == "__main__":
    generate_excel("receipts.json", "depenses.xlsx")

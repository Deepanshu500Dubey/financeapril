"""
forecast_editor.py
==================
Natural Language Forecast Editor

Allows editing of the Q3 2026 forecast models via plain-English prompts such as:
  - "Increase the forecast by 10%"
  - "Reduce the sales forecast for Q2 by 5%"
  - "Increase the revenue forecast for Product A by 20%"

The editor:
  1. Maintains driver state for both the quarterly (build_forecast.py) and
     driver-based (build_driver_forecast.py) forecast models.
  2. Parses NL prompts to identify target field, scenario, period, and magnitude.
  3. Applies the delta and records a detailed change log.
  4. Generates modified Excel bytes (in-memory execution of patched build scripts).
  5. Produces a comparison Excel report highlighting all changes.
"""

from __future__ import annotations

import copy
import io
import os
import re
import sys
import tempfile
import subprocess
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── COLOUR PALETTE FOR COMPARISON REPORT ────────────────────────────────────
_NAVY   = "1F3864"
_MID    = "2E75B6"
_LIGHT  = "BDD7EE"
_GREEN  = "E2EFDA"
_RED    = "FFE5E5"
_AMBER  = "FFF2CC"
_WHITE  = "FFFFFF"
_GRAY   = "F2F2F2"
_DKGRAY = "404040"
_CHNG_OLD = "FFDAC1"   # pastel orange  – original value (changed)
_CHNG_NEW = "C6EFCE"   # pastel green   – new value (changed)

SCRIPT_BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ─── CHANGE LOG ENTRY ─────────────────────────────────────────────────────────

@dataclass
class ChangeEntry:
    model: str          # "quarterly" | "driver"
    driver_key: str     # canonical key name
    label: str          # human-readable field label
    scenario: str       # "base" | "upside" | "risk" | "all"
    old_value: Any
    new_value: Any
    delta_pct: Optional[float]
    prompt: str         # original NL prompt that caused this change
    applied_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        old_fmt = _fmt(self.old_value)
        new_fmt = _fmt(self.new_value)
        delta = (self.new_value - self.old_value) if isinstance(self.new_value, (int, float)) else None
        delta_pct_disp = f"{self.delta_pct:+.1f}%" if self.delta_pct is not None else "N/A"
        return {
            "model": self.model,
            "driver_key": self.driver_key,
            "label": self.label,
            "scenario": self.scenario,
            "old_value": self.old_value,
            "new_value": self.new_value,
            "old_value_formatted": old_fmt,
            "new_value_formatted": new_fmt,
            "delta": delta,
            "delta_pct": self.delta_pct,
            "delta_pct_formatted": delta_pct_disp,
            "prompt": self.prompt,
            "applied_at": self.applied_at,
        }


def _fmt(val: Any) -> str:
    if isinstance(val, float):
        if 0 < abs(val) < 2:  # likely a percentage
            return f"{val:.2%}"
        return f"{val:,.2f}"
    if isinstance(val, int):
        return f"{val:,}"
    return str(val)


# ─── DRIVER STATE DEFINITIONS ─────────────────────────────────────────────────

def _quarterly_defaults() -> Dict[str, Any]:
    """
    Mirrors the ACT dict and sections list in build_forecast.py.
    Structure: {driver_key: {"label": ..., "base": val, "upside": val, "risk": val, "type": "pct"|"dollar"|"int"}}
    For simple scalars (ACT dict): store under "base" only.
    """
    return {
        # ── REVENUE ──────────────────────────────────────────────────────────
        "revenue_baseline":      {"label": "May 2026 Baseline Revenue (AUD)",         "base": 40546176, "upside": 40546176, "risk": 40546176, "type": "dollar"},
        "revenue_growth_rate":   {"label": "Revenue Growth Rate (quarterly)",          "base": 0.05,     "upside": 0.20,     "risk": 0.10,     "type": "pct"},
        "product_mix":           {"label": "Product Sales Mix (%)",                    "base": 0.737,    "upside": 0.737,    "risk": 0.737,    "type": "pct"},
        "service_mix":           {"label": "Service Revenue Mix (%)",                  "base": 0.156,    "upside": 0.156,    "risk": 0.156,    "type": "pct"},
        "subscription_mix":      {"label": "Subscription Revenue Mix (%)",             "base": 0.086,    "upside": 0.086,    "risk": 0.086,    "type": "pct"},

        # ── GEOPOLITICAL COST OVERLAYS ────────────────────────────────────────
        "logistics_uplift":      {"label": "Logistics / Courier uplift (%)",           "base": 0.05,     "upside": 0.05,     "risk": 0.15,     "type": "pct"},
        "insurance_uplift":      {"label": "Insurance (Marine/Cargo) uplift (%)",      "base": 0.05,     "upside": 0.05,     "risk": 0.12,     "type": "pct"},
        "fuel_uplift":           {"label": "Fuel-linked costs (Telecom+Utilities) uplift (%)","base": 0.03,"upside": 0.03,"risk": 0.10,       "type": "pct"},
        "supplier_pricing":      {"label": "Supplier pricing pressure (COGS proxy) (%)", "base": 0.02,   "upside": 0.02,    "risk": 0.10,     "type": "pct"},
        "fx_loss_uplift":        {"label": "FX Loss uplift (%)",                       "base": 0.05,     "upside": 0.05,     "risk": 0.20,     "type": "pct"},

        # ── PERSONNEL ─────────────────────────────────────────────────────────
        "salary_growth":         {"label": "Salary growth rate (%)",                   "base": 0.02,     "upside": 0.025,    "risk": 0.02,     "type": "pct"},
        "benefits_growth":       {"label": "Benefits growth rate (%)",                 "base": 0.02,     "upside": 0.025,    "risk": 0.02,     "type": "pct"},
        "contractor_growth":     {"label": "Contractor growth rate (%)",               "base": 0.01,     "upside": 0.030,    "risk": 0.01,     "type": "pct"},
        "marketing_uplift":      {"label": "Marketing uplift (%)",                     "base": 0.03,     "upside": 0.08,     "risk": 0.03,     "type": "pct"},
        "te_uplift":             {"label": "T&E uplift (%)",                           "base": 0.02,     "upside": 0.04,     "risk": 0.02,     "type": "pct"},

        # ── WORKING CAPITAL ───────────────────────────────────────────────────
        "dso_days":              {"label": "Debtor Days (DSO)",                        "base": 45,       "upside": 45,       "risk": 55,       "type": "int"},
        "inventory_buffer":      {"label": "Inventory buffer build (AUD)",             "base": 0,        "upside": 0,        "risk": 500000,   "type": "dollar"},
        "opening_cash":          {"label": "Opening Cash Balance (AUD)",               "base": 12911223, "upside": 12911223, "risk": 12911223, "type": "dollar"},
        "capex_q3":              {"label": "Capex Q3 forecast (AUD)",                  "base": 1500000,  "upside": 2000000,  "risk": 1200000,  "type": "dollar"},
        "tax_rate":              {"label": "Tax rate (%)",                              "base": 0.30,     "upside": 0.30,     "risk": 0.30,     "type": "pct"},
        "min_cash_buffer":       {"label": "Min cash buffer target (AUD)",             "base": 5000000,  "upside": 5000000,  "risk": 7000000,  "type": "dollar"},
    }


def _driver_defaults() -> Dict[str, Any]:
    """
    Mirrors the drow() calls in build_driver_forecast.py.
    Base/Upside/Risk match the first three numeric args of each drow() call.
    """
    return {
        # ── REVENUE DRIVERS ───────────────────────────────────────────────────
        "revenue_baseline":      {"label": "May 2026 Monthly Revenue Baseline ($)",    "base": 40546176, "upside": 40546176, "risk": 40546176, "type": "dollar"},
        "revenue_growth_rate":   {"label": "YoY Revenue Growth Rate",                  "base": 0.088,    "upside": 0.20,     "risk": 0.045,    "type": "pct"},
        "jul_seasonality":       {"label": "Aug Seasonality Factor",                   "base": 0.95,     "upside": 0.95,     "risk": 0.90,     "type": "pct"},
        "aug_seasonality":       {"label": "Aug Seasonality Factor",                   "base": 1.00,     "upside": 1.05,     "risk": 0.95,     "type": "pct"},
        "sep_seasonality":       {"label": "Sep Seasonality Factor",                   "base": 1.08,     "upside": 1.20,     "risk": 1.00,     "type": "pct"},
        "product_mix":           {"label": "Product Revenue Mix (%)",                  "base": 0.415,    "upside": 0.415,    "risk": 0.40,     "type": "pct"},
        "service_mix":           {"label": "Service Revenue Mix (%)",                  "base": 0.385,    "upside": 0.385,    "risk": 0.375,    "type": "pct"},
        "subscription_mix":      {"label": "Subscription Revenue Mix (%)",             "base": 0.20,     "upside": 0.20,     "risk": 0.225,    "type": "pct"},

        # ── FIXED COST DRIVERS ────────────────────────────────────────────────
        "workforce_pct":         {"label": "Workforce Cost as % Revenue",              "base": 0.084,    "upside": 0.080,    "risk": 0.090,    "type": "pct"},
        "prepay_amort":          {"label": "Prepayment Amortization (monthly $)",      "base": 431000,   "upside": 431000,   "risk": 431000,   "type": "dollar"},
        "depreciation":          {"label": "Depreciation (monthly $)",                 "base": 650000,   "upside": 650000,   "risk": 650000,   "type": "dollar"},
        "insurance_pct":         {"label": "Insurance % Revenue",                      "base": 0.022,    "upside": 0.022,    "risk": 0.028,    "type": "pct"},
        "software_licences":     {"label": "Software/Licences (monthly $)",            "base": 380000,   "upside": 380000,   "risk": 400000,   "type": "dollar"},

        # ── VARIABLE COST DRIVERS ─────────────────────────────────────────────
        "freight_pct":           {"label": "Postage & Freight % Revenue",              "base": 0.048,    "upside": 0.048,    "risk": 0.062,    "type": "pct"},
        "telecom_pct":           {"label": "Telecom & Utilities % Revenue",            "base": 0.019,    "upside": 0.019,    "risk": 0.023,    "type": "pct"},
        "advertising_pct":       {"label": "Advertising % Revenue",                   "base": 0.036,    "upside": 0.040,    "risk": 0.034,    "type": "pct"},
        "contractor_pct":        {"label": "Contractor/Consulting % Revenue",         "base": 0.028,    "upside": 0.030,    "risk": 0.025,    "type": "pct"},
        "supplier_uplift":       {"label": "Supplier Pricing Uplift (geopolitical)",  "base": 0.00,     "upside": 0.00,     "risk": 0.035,    "type": "pct"},
        "other_opex_pct":        {"label": "Other OpEx % Revenue",                    "base": 0.055,    "upside": 0.052,    "risk": 0.060,    "type": "pct"},

        # ── WORKING CAPITAL DRIVERS ───────────────────────────────────────────
        "dso_days":              {"label": "Days Sales Outstanding (DSO)",             "base": 45,       "upside": 42,       "risk": 55,       "type": "int"},
        "ar_collection_current": {"label": "AR Collection Rate — Current (<30d)",      "base": 0.95,     "upside": 0.97,     "risk": 0.88,     "type": "pct"},
        "ar_collection_30_60":   {"label": "AR Collection Rate — 30-60d",              "base": 0.75,     "upside": 0.80,     "risk": 0.60,     "type": "pct"},
        "ar_collection_60_90":   {"label": "AR Collection Rate — 60-90d",              "base": 0.55,     "upside": 0.60,     "risk": 0.40,     "type": "pct"},
        "ar_collection_90plus":  {"label": "AR Collection Rate — >90d",                "base": 0.30,     "upside": 0.35,     "risk": 0.15,     "type": "pct"},
        "dpo_days":              {"label": "Days Payable Outstanding (DPO)",           "base": 38,       "upside": 38,       "risk": 42,       "type": "int"},
        "monthly_capex":         {"label": "Monthly Capex ($)",                        "base": 400000,   "upside": 600000,   "risk": 300000,   "type": "dollar"},

        # ── TAX & INTEREST ────────────────────────────────────────────────────
        "tax_rate":              {"label": "Effective Tax Rate",                       "base": 0.30,     "upside": 0.30,     "risk": 0.30,     "type": "pct"},
        "interest_rate":         {"label": "Interest Income % Cash",                  "base": 0.0045,   "upside": 0.0045,   "risk": 0.0035,   "type": "pct"},
        "fx_impact":             {"label": "FX Impact on Revenue",                     "base": 0.008,    "upside": 0.005,    "risk": 0.015,    "type": "pct"},
    }


# ─── NL INTENT MAPPING ────────────────────────────────────────────────────────

# Maps keywords → list of matching driver keys (in priority order)
_QUARTERLY_KEYWORD_MAP: List[Tuple[List[str], List[str]]] = [
    # (keyword list, [driver_keys_for_quarterly, driver_keys_for_driver])
    (["overall revenue", "total revenue", "forecast", "revenue"],            ["revenue_baseline", "revenue_growth_rate"]),
    (["growth rate", "yoy", "year-on-year", "growth"],                        ["revenue_growth_rate"]),
    (["product", "product sales", "product revenue", "product a"],            ["product_mix"]),
    (["service", "service revenue"],                                          ["service_mix"]),
    (["subscription", "recurring"],                                           ["subscription_mix"]),
    (["logistics", "freight", "postage", "courier", "shipping"],              ["logistics_uplift"]),
    (["insurance"],                                                            ["insurance_uplift"]),
    (["fuel", "utilities", "telecom", "energy"],                              ["fuel_uplift"]),
    (["supplier", "cogs", "cost of goods", "input cost"],                     ["supplier_pricing"]),
    (["fx", "foreign exchange", "currency"],                                  ["fx_loss_uplift"]),
    (["salary", "salaries", "wages", "personnel", "workforce", "headcount"], ["salary_growth"]),
    (["benefits", "employee benefits"],                                        ["benefits_growth"]),
    (["contractor", "consulting", "consultant"],                              ["contractor_growth"]),
    (["marketing", "advertising", "promotion"],                               ["marketing_uplift"]),
    (["travel", "t&e", "entertainment"],                                      ["te_uplift"]),
    (["dso", "debtor", "receivables", "days sales outstanding"],              ["dso_days"]),
    (["capex", "capital expenditure", "investment"],                          ["capex_q3"]),
    (["tax", "tax rate"],                                                     ["tax_rate"]),
    (["cash", "cash balance", "opening cash"],                                ["opening_cash"]),
]

_DRIVER_KEYWORD_MAP: List[Tuple[List[str], List[str]]] = [
    (["overall revenue", "total revenue", "forecast", "revenue"],             ["revenue_baseline", "revenue_growth_rate"]),
    (["growth rate", "yoy", "year-on-year", "growth"],                        ["revenue_growth_rate"]),
    (["product", "product sales", "product revenue", "product a"],            ["product_mix"]),
    (["service", "service revenue"],                                          ["service_mix"]),
    (["subscription", "recurring"],                                           ["subscription_mix"]),
    (["jul", "july", "q3 jul", "q3-jul"],                                    ["jul_seasonality"]),
    (["aug", "august", "q3 aug"],                                             ["aug_seasonality"]),
    (["sep", "september", "q3 sep"],                                          ["sep_seasonality"]),
    # Q2 maps to Aug (nearest mid-quarter)
    (["q2", "second quarter"],                                                ["aug_seasonality"]),
    (["logistics", "freight", "postage", "courier", "shipping"],              ["freight_pct"]),
    (["insurance"],                                                            ["insurance_pct"]),
    (["telecom", "utilities", "fuel", "energy"],                              ["telecom_pct"]),
    (["advertising", "marketing", "promotion"],                               ["advertising_pct"]),
    (["contractor", "consulting", "consultant"],                              ["contractor_pct"]),
    (["supplier", "geopolitical", "input cost"],                              ["supplier_uplift"]),
    (["other opex", "other operating", "opex"],                               ["other_opex_pct"]),
    (["salary", "salaries", "wages", "personnel", "workforce", "headcount"], ["workforce_pct"]),
    (["depreciation", "d&a", "amortisation", "amortization"],                ["depreciation"]),
    (["software", "licences", "licenses"],                                    ["software_licences"]),
    (["prepay", "prepayment"],                                                ["prepay_amort"]),
    (["dso", "debtor", "receivables", "days sales outstanding"],              ["dso_days"]),
    (["dpo", "payables", "days payable"],                                     ["dpo_days"]),
    (["capex", "capital expenditure", "investment"],                          ["monthly_capex"]),
    (["tax", "tax rate"],                                                     ["tax_rate"]),
    (["interest", "interest income"],                                         ["interest_rate"]),
    (["fx", "foreign exchange", "currency"],                                  ["fx_impact"]),
    (["ar collection", "collection rate"],                                    ["ar_collection_current", "ar_collection_30_60"]),
]


# ─── INTENT PARSER ────────────────────────────────────────────────────────────

def _extract_magnitude(prompt: str) -> Tuple[Optional[float], str]:
    """
    Returns (pct_delta, mode) where:
      pct_delta  = float percentage change (e.g. +10.0, -5.0)
      mode       = "pct_change" | "pct_set" | "unknown"

    Examples:
      "increase by 10%" → (+10.0, "pct_change")
      "reduce by 5%"    → (-5.0,  "pct_change")
      "set to 15%"      → (+15.0, "pct_set")
    """
    p = prompt.lower().strip()

    # Detect direction
    is_increase = bool(re.search(r'\b(increas|grow|uplift|boost|rais|add)\w*', p))
    is_decrease = bool(re.search(r'\b(decreas|reduc|lower|cut|drop|shrink|diminish)\w*', p))

    # Extract numeric value + optional %
    match = re.search(r'(\d+(?:\.\d+)?)\s*%?', p)
    if not match:
        return None, "unknown"

    val = float(match.group(1))

    # Determine if it's a relative change or absolute set
    if re.search(r'\b(set to|change to|update to|make it|adjust to)\b', p):
        return val, "pct_set"

    if is_decrease:
        return -val, "pct_change"
    if is_increase:
        return +val, "pct_change"

    # Default: treat bare "by X%" as increase
    return +val, "pct_change"


def _detect_scenario(prompt: str) -> List[str]:
    """
    Returns list of scenario keys to modify. Defaults to ['base'].
    """
    p = prompt.lower()
    scenarios = []
    if re.search(r'\b(upside|optimist|bullish|best case|best-case)\b', p):
        scenarios.append("upside")
    if re.search(r'\b(risk|geopolit|pessimist|bearish|worst case|worst-case|downside)\b', p):
        scenarios.append("risk")
    if re.search(r'\b(all scenario|all three|every scenario)\b', p):
        return ["base", "upside", "risk"]
    if not scenarios:
        scenarios = ["base"]  # default
    return scenarios


def _detect_target_keys(prompt: str, keyword_map: List[Tuple[List[str], List[str]]]) -> List[str]:
    """
    Match prompt against keyword map; return driver keys in order of confidence.
    Longer keyword matches beat shorter ones.
    """
    p = prompt.lower().strip()
    best_keys: List[str] = []
    best_len = 0

    for keywords, driver_keys in keyword_map:
        for kw in keywords:
            if kw in p and len(kw) > best_len:
                best_len = len(kw)
                best_keys = driver_keys

    # Universal fallback: any prompt with increase/decrease and no specific match
    # defaults to revenue_baseline (the primary driver people think of as "the forecast")
    if not best_keys:
        best_keys = ["revenue_baseline"]

    return best_keys


def _parse_debug(prompt: str) -> Dict[str, Any]:
    """Return a full debug breakdown of how a prompt is parsed."""
    p = prompt.lower().strip()
    pct_delta, mode = _extract_magnitude(prompt)
    scenarios = _detect_scenario(prompt)
    keys_q = _detect_target_keys(prompt, _QUARTERLY_KEYWORD_MAP)
    keys_d = _detect_target_keys(prompt, _DRIVER_KEYWORD_MAP)
    return {
        "prompt": prompt,
        "prompt_lowercased": p,
        "pct_delta": pct_delta,
        "mode": mode,
        "scenarios": scenarios,
        "quarterly_target_keys": keys_q,
        "driver_target_keys": keys_d,
        "keyword_checks": {
            "forecast_in_prompt": "forecast" in p,
            "revenue_in_prompt": "revenue" in p,
            "increase_detected": bool(re.search(r'\b(increas|grow|uplift|boost|rais|add)\w*', p)),
            "decrease_detected": bool(re.search(r'\b(decreas|reduc|lower|cut|drop|shrink|diminish)\w*', p)),
        },
    }


# ─── MAIN EDITOR CLASS ────────────────────────────────────────────────────────

class ForecastEditor:
    """
    Stateful forecast editor. Maintains separate states for 'quarterly' and 'driver' models.
    Thread-safety note: not thread-safe by default; use one instance per session/request context
    or add locking for concurrent use.
    """

    def __init__(self):
        self._quarterly_state: Dict[str, Any] = _quarterly_defaults()
        self._driver_state: Dict[str, Any] = _driver_defaults()
        self._quarterly_original: Dict[str, Any] = copy.deepcopy(self._quarterly_state)
        self._driver_original: Dict[str, Any] = copy.deepcopy(self._driver_state)
        self.change_log: List[ChangeEntry] = []

    # ── Public API ────────────────────────────────────────────────────────────

    def parse_and_apply(
        self,
        prompt: str,
        model: str = "both",       # "quarterly" | "driver" | "both"
        scenario_override: Optional[str] = None,   # override auto-detection
    ) -> List[Dict[str, Any]]:
        """
        Parse an NL prompt and apply changes to the state.
        Returns the list of changes made (as dicts).
        """
        pct_delta, mode = _extract_magnitude(prompt)
        if pct_delta is None:
            raise ValueError(
                f"Could not extract a magnitude from prompt: '{prompt}'. "
                "Please include a percentage value, e.g. 'Increase by 10%'."
            )

        scenarios = [scenario_override] if scenario_override else _detect_scenario(prompt)

        changes_made: List[Dict[str, Any]] = []
        models_to_edit = ["quarterly", "driver"] if model == "both" else [model]

        for mdl in models_to_edit:
            state = self._quarterly_state if mdl == "quarterly" else self._driver_state
            kmap = _QUARTERLY_KEYWORD_MAP if mdl == "quarterly" else _DRIVER_KEYWORD_MAP
            target_keys = _detect_target_keys(prompt, kmap)

            if not target_keys:
                continue

            for key in target_keys:
                if key not in state:
                    continue
                entry = state[key]
                for sc in scenarios:
                    if sc not in entry:
                        continue
                    old_val = entry[sc]
                    new_val = _apply_delta(old_val, pct_delta, mode, entry["type"])
                    if new_val == old_val:
                        continue

                    entry[sc] = new_val

                    change = ChangeEntry(
                        model=mdl,
                        driver_key=key,
                        label=entry["label"],
                        scenario=sc,
                        old_value=old_val,
                        new_value=new_val,
                        delta_pct=pct_delta,
                        prompt=prompt,
                    )
                    self.change_log.append(change)
                    changes_made.append(change.to_dict())

        if not changes_made:
            raise ValueError(
                f"No matching driver found for prompt: '{prompt}'. "
                "Try keywords like 'revenue', 'sales', 'logistics', 'insurance', 'workforce', etc."
            )

        return changes_made

    def reset(self, model: str = "both"):
        """Reset state back to original values."""
        if model in ("quarterly", "both"):
            self._quarterly_state = copy.deepcopy(self._quarterly_original)
        if model in ("driver", "both"):
            self._driver_state = copy.deepcopy(self._driver_original)
        self.change_log.clear()

    def get_state(self, model: str = "both") -> Dict[str, Any]:
        """Return current driver state as a serialisable dict."""
        if model == "quarterly":
            return {"quarterly": self._quarterly_state}
        if model == "driver":
            return {"driver": self._driver_state}
        return {"quarterly": self._quarterly_state, "driver": self._driver_state}

    def get_change_log(self) -> List[Dict[str, Any]]:
        return [c.to_dict() for c in self.change_log]

    # ── Excel Generation ──────────────────────────────────────────────────────

    def generate_quarterly_excel_bytes(self) -> bytes:
        """Run build_forecast.py with patched values and return xlsx bytes."""
        script_path = os.path.join(SCRIPT_BASE_DIR, "build_forecast.py")
        code = _read_file(script_path)
        code = _patch_quarterly_script(code, self._quarterly_state)
        return _run_patched_script(code, suffix=".xlsx")

    def generate_driver_excel_bytes(self) -> bytes:
        """Run build_driver_forecast.py with patched values and return xlsx bytes."""
        script_path = os.path.join(SCRIPT_BASE_DIR, "build_driver_forecast.py")
        code = _read_file(script_path)
        code = _patch_driver_script(code, self._driver_state)
        return _run_patched_script(code, suffix=".xlsx")

    def generate_comparison_excel_bytes(self, model: str = "both") -> bytes:
        """
        Build a comparison Excel report that shows original vs. modified values
        for all changed drivers.
        """
        wb = Workbook()
        ws = wb.active
        ws.title = "Forecast Changes"

        _build_comparison_sheet(ws, self.change_log, model)

        # Add summary sheet
        ws_summary = wb.create_sheet("Summary")
        _build_summary_sheet(ws_summary, self.change_log)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()


# ─── SCRIPT PATCHING HELPERS ──────────────────────────────────────────────────

def _read_file(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def _run_patched_script(code: str, suffix: str = ".xlsx") -> bytes:
    """Write patched code to a temp file, execute it, return output file bytes."""
    tmp_out = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    tmp_out.close()
    out_path = tmp_out.name

    # Replace save paths
    code = code.replace(
        '"/sessions/practical-blissful-maxwell/mnt/outputs/Q3_2026_Quarterly_Forecast.xlsx"',
        repr(out_path.replace("\\", "/"))
    )
    code = code.replace(
        '"/sessions/practical-blissful-maxwell/mnt/outputs/Driver_Forecast_Q3_2026.xlsx"',
        repr(out_path.replace("\\", "/"))
    )
    code = code.replace(
        '"/sessions/practical-blissful-maxwell/mnt/outputs/Driver_Forecast_Q3_2026.xlsm"',
        repr(out_path.replace("\\", "/"))
    )
    # Also replace OUT_XLSX / OUT_XLSM assignments
    code = re.sub(
        r'OUT_XLSX\s*=\s*"[^"]+"',
        f'OUT_XLSX = {repr(out_path.replace(chr(92), "/"))}',
        code
    )
    code = re.sub(
        r'OUT_XLSM\s*=\s*"[^"]+"',
        f'OUT_XLSM = {repr(out_path.replace(chr(92), "/"))}',
        code
    )

    tmp_script = tempfile.NamedTemporaryFile(mode="w", suffix=".py", delete=False, encoding="utf-8")
    tmp_script.write(code)
    tmp_script.close()

    try:
        result = subprocess.run(
            [sys.executable, tmp_script.name],
            capture_output=True,
            text=True,
            timeout=180,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("Script execution timed out (180s)")
    finally:
        try:
            os.unlink(tmp_script.name)
        except OSError:
            pass

    if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
        detail = result.stderr[-2000:] if result.stderr else "Unknown error"
        raise RuntimeError(f"Script execution failed: {detail}")

    with open(out_path, "rb") as f:
        content = f.read()
    try:
        os.unlink(out_path)
    except OSError:
        pass

    return content


def _patch_quarterly_script(code: str, state: Dict[str, Any]) -> str:
    """
    Patch the quarterly forecast script (build_forecast.py) with modified driver values.
    The script stores values directly in the ACT dict and sections list.
    We use targeted string replacements on the numeric literals.
    """
    s = state

    # ── ACT dict values ──────────────────────────────────────────────────────
    # ACT["total_rev"] — revenue baseline
    code = _replace_act_value(code, "total_rev", s["revenue_baseline"]["base"])
    # Note: product_sales, service_rev, sub_rev are computed from total_rev * mix,
    # so we patch those proportionally.
    total_rev = s["revenue_baseline"]["base"]
    code = _replace_act_value(code, "product_sales", total_rev * s["product_mix"]["base"])
    code = _replace_act_value(code, "service_rev",   total_rev * s["service_mix"]["base"])
    code = _replace_act_value(code, "sub_rev",       total_rev * s["subscription_mix"]["base"])

    # ── sections list (Assumptions sheet) ────────────────────────────────────
    code = _replace_assumption(code, "May 2026 Baseline Revenue (AUD)",
                                s["revenue_baseline"]["base"],
                                s["revenue_baseline"]["upside"],
                                s["revenue_baseline"]["risk"])
    code = _replace_assumption(code, "Revenue Growth Rate (quarterly)",
                                s["revenue_growth_rate"]["base"],
                                s["revenue_growth_rate"]["upside"],
                                s["revenue_growth_rate"]["risk"])
    code = _replace_assumption(code, "Product Sales Mix (%)",
                                s["product_mix"]["base"],
                                s["product_mix"]["upside"],
                                s["product_mix"]["risk"])
    code = _replace_assumption(code, "Service Revenue Mix (%)",
                                s["service_mix"]["base"],
                                s["service_mix"]["upside"],
                                s["service_mix"]["risk"])
    code = _replace_assumption(code, "Subscription Revenue Mix (%)",
                                s["subscription_mix"]["base"],
                                s["subscription_mix"]["upside"],
                                s["subscription_mix"]["risk"])
    code = _replace_assumption(code, "Logistics / Courier uplift (%)",
                                s["logistics_uplift"]["base"],
                                s["logistics_uplift"]["upside"],
                                s["logistics_uplift"]["risk"])
    code = _replace_assumption(code, "Insurance (Marine/Cargo) uplift (%)",
                                s["insurance_uplift"]["base"],
                                s["insurance_uplift"]["upside"],
                                s["insurance_uplift"]["risk"])
    code = _replace_assumption(code, "Fuel-linked costs (Telecom+Utilities) uplift (%)",
                                s["fuel_uplift"]["base"],
                                s["fuel_uplift"]["upside"],
                                s["fuel_uplift"]["risk"])
    code = _replace_assumption(code, "Supplier pricing pressure (COGS proxy) (%)",
                                s["supplier_pricing"]["base"],
                                s["supplier_pricing"]["upside"],
                                s["supplier_pricing"]["risk"])
    code = _replace_assumption(code, "FX Loss uplift (%)",
                                s["fx_loss_uplift"]["base"],
                                s["fx_loss_uplift"]["upside"],
                                s["fx_loss_uplift"]["risk"])
    code = _replace_assumption(code, "Salary growth rate (%)",
                                s["salary_growth"]["base"],
                                s["salary_growth"]["upside"],
                                s["salary_growth"]["risk"])
    code = _replace_assumption(code, "Benefits growth rate (%)",
                                s["benefits_growth"]["base"],
                                s["benefits_growth"]["upside"],
                                s["benefits_growth"]["risk"])
    code = _replace_assumption(code, "Contractor growth rate (%)",
                                s["contractor_growth"]["base"],
                                s["contractor_growth"]["upside"],
                                s["contractor_growth"]["risk"])
    code = _replace_assumption(code, "Marketing uplift (%)",
                                s["marketing_uplift"]["base"],
                                s["marketing_uplift"]["upside"],
                                s["marketing_uplift"]["risk"])
    code = _replace_assumption(code, "T&E uplift (%)",
                                s["te_uplift"]["base"],
                                s["te_uplift"]["upside"],
                                s["te_uplift"]["risk"])
    code = _replace_assumption(code, "Debtor Days (DSO) - days",
                                s["dso_days"]["base"],
                                s["dso_days"]["upside"],
                                s["dso_days"]["risk"])
    code = _replace_assumption(code, "Opening Cash Balance (AUD)",
                                s["opening_cash"]["base"],
                                s["opening_cash"]["upside"],
                                s["opening_cash"]["risk"])
    code = _replace_assumption(code, "Capex Q3 forecast (AUD)",
                                s["capex_q3"]["base"],
                                s["capex_q3"]["upside"],
                                s["capex_q3"]["risk"])
    code = _replace_assumption(code, "Tax rate (%)",
                                s["tax_rate"]["base"],
                                s["tax_rate"]["upside"],
                                s["tax_rate"]["risk"])
    code = _replace_assumption(code, "Min cash buffer target (AUD)",
                                s["min_cash_buffer"]["base"],
                                s["min_cash_buffer"]["upside"],
                                s["min_cash_buffer"]["risk"])

    return code


def _patch_driver_script(code: str, state: Dict[str, Any]) -> str:
    """
    Patch the driver forecast script (build_driver_forecast.py) with modified driver values.
    Uses targeted replacements on drow() calls.
    """
    s = state
    code = _replace_drow(code, "May 2026 Monthly Revenue Baseline ($)",
                          s["revenue_baseline"]["base"],
                          s["revenue_baseline"]["upside"],
                          s["revenue_baseline"]["risk"])
    code = _replace_drow(code, "YoY Revenue Growth Rate",
                          s["revenue_growth_rate"]["base"],
                          s["revenue_growth_rate"]["upside"],
                          s["revenue_growth_rate"]["risk"])
    code = _replace_drow(code, "Aug Seasonality Factor",
                          s["jul_seasonality"]["base"],
                          s["jul_seasonality"]["upside"],
                          s["jul_seasonality"]["risk"])
    code = _replace_drow(code, "Aug Seasonality Factor",
                          s["aug_seasonality"]["base"],
                          s["aug_seasonality"]["upside"],
                          s["aug_seasonality"]["risk"])
    code = _replace_drow(code, "Sep Seasonality Factor",
                          s["sep_seasonality"]["base"],
                          s["sep_seasonality"]["upside"],
                          s["sep_seasonality"]["risk"])
    code = _replace_drow(code, "Product Revenue Mix (%)",
                          s["product_mix"]["base"],
                          s["product_mix"]["upside"],
                          s["product_mix"]["risk"])
    code = _replace_drow(code, "Service Revenue Mix (%)",
                          s["service_mix"]["base"],
                          s["service_mix"]["upside"],
                          s["service_mix"]["risk"])
    code = _replace_drow(code, "Subscription Revenue Mix (%)",
                          s["subscription_mix"]["base"],
                          s["subscription_mix"]["upside"],
                          s["subscription_mix"]["risk"])
    code = _replace_drow(code, "Workforce Cost as % Revenue",
                          s["workforce_pct"]["base"],
                          s["workforce_pct"]["upside"],
                          s["workforce_pct"]["risk"])
    code = _replace_drow(code, "Prepayment Amortization (monthly $)",
                          s["prepay_amort"]["base"],
                          s["prepay_amort"]["upside"],
                          s["prepay_amort"]["risk"])
    code = _replace_drow(code, "Depreciation (monthly $)",
                          s["depreciation"]["base"],
                          s["depreciation"]["upside"],
                          s["depreciation"]["risk"])
    code = _replace_drow(code, "Insurance % Revenue",
                          s["insurance_pct"]["base"],
                          s["insurance_pct"]["upside"],
                          s["insurance_pct"]["risk"])
    code = _replace_drow(code, "Postage & Freight % Revenue",
                          s["freight_pct"]["base"],
                          s["freight_pct"]["upside"],
                          s["freight_pct"]["risk"])
    code = _replace_drow(code, "Telecom & Utilities % Revenue",
                          s["telecom_pct"]["base"],
                          s["telecom_pct"]["upside"],
                          s["telecom_pct"]["risk"])
    code = _replace_drow(code, "Advertising % Revenue",
                          s["advertising_pct"]["base"],
                          s["advertising_pct"]["upside"],
                          s["advertising_pct"]["risk"])
    code = _replace_drow(code, "Contractor/Consulting % Revenue",
                          s["contractor_pct"]["base"],
                          s["contractor_pct"]["upside"],
                          s["contractor_pct"]["risk"])
    code = _replace_drow(code, "Supplier Pricing Uplift (geopolitical)",
                          s["supplier_uplift"]["base"],
                          s["supplier_uplift"]["upside"],
                          s["supplier_uplift"]["risk"])
    code = _replace_drow(code, "Other OpEx % Revenue",
                          s["other_opex_pct"]["base"],
                          s["other_opex_pct"]["upside"],
                          s["other_opex_pct"]["risk"])
    code = _replace_drow(code, "Days Sales Outstanding (DSO)",
                          s["dso_days"]["base"],
                          s["dso_days"]["upside"],
                          s["dso_days"]["risk"])
    code = _replace_drow(code, "AR Collection Rate \u2014 Current (<30d)",
                          s["ar_collection_current"]["base"],
                          s["ar_collection_current"]["upside"],
                          s["ar_collection_current"]["risk"])
    code = _replace_drow(code, "AR Collection Rate \u2014 30-60d",
                          s["ar_collection_30_60"]["base"],
                          s["ar_collection_30_60"]["upside"],
                          s["ar_collection_30_60"]["risk"])
    code = _replace_drow(code, "AR Collection Rate \u2014 60-90d",
                          s["ar_collection_60_90"]["base"],
                          s["ar_collection_60_90"]["upside"],
                          s["ar_collection_60_90"]["risk"])
    code = _replace_drow(code, "AR Collection Rate \u2014 >90d",
                          s["ar_collection_90plus"]["base"],
                          s["ar_collection_90plus"]["upside"],
                          s["ar_collection_90plus"]["risk"])
    code = _replace_drow(code, "Days Payable Outstanding (DPO)",
                          s["dpo_days"]["base"],
                          s["dpo_days"]["upside"],
                          s["dpo_days"]["risk"])
    code = _replace_drow(code, "Monthly Capex ($)",
                          s["monthly_capex"]["base"],
                          s["monthly_capex"]["upside"],
                          s["monthly_capex"]["risk"])
    code = _replace_drow(code, "Effective Tax Rate",
                          s["tax_rate"]["base"],
                          s["tax_rate"]["upside"],
                          s["tax_rate"]["risk"])
    code = _replace_drow(code, "Interest Income % Cash",
                          s["interest_rate"]["base"],
                          s["interest_rate"]["upside"],
                          s["interest_rate"]["risk"])
    code = _replace_drow(code, "FX Impact on Revenue",
                          s["fx_impact"]["base"],
                          s["fx_impact"]["upside"],
                          s["fx_impact"]["risk"])
    return code


def _num_repr(val: Any) -> str:
    """Format number for Python source code replacement."""
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        # Use up to 6 significant digits
        return f"{val:.6g}"
    return repr(val)


def _replace_act_value(code: str, key: str, new_val: float) -> str:
    """Replace a value in the ACT dict."""
    # Match: "key":    12345.67,
    pattern = rf'("{re.escape(key)}":\s*)([\d.]+)(,)'
    replacement = lambda m: m.group(1) + _num_repr(new_val) + m.group(3)
    new_code, n = re.subn(pattern, replacement, code, count=1)
    return new_code


def _replace_assumption(code: str, label: str, base: Any, upside: Any, risk: Any) -> str:
    """
    Replace the numeric triple (base, upside, risk) in a sections-list tuple.
    Matches lines like:  ("Label text",0.00,0.20,0.10,"Notes"),
    """
    esc = re.escape(label)
    # Loosely match: ("label", num1, num2, num3, "notes")
    pattern = rf'("{esc}"\s*,\s*)([\d.]+)(\s*,\s*)([\d.]+)(\s*,\s*)([\d.]+)'

    def repl(m):
        return (m.group(1) +
                _num_repr(base) + m.group(3) +
                _num_repr(upside) + m.group(5) +
                _num_repr(risk))

    new_code, n = re.subn(pattern, repl, code, count=1)
    return new_code


def _replace_drow(code: str, label: str, base: Any, upside: Any, risk: Any) -> str:
    """
    Replace the numeric triple (base, upside, risk) in a drow() call.
    Matches: drow(ws,r,"Label text",base,upside,risk,FMT,"notes");
    """
    esc = re.escape(label)
    pattern = rf'(drow\(ws,r,"{esc}"\s*,\s*)([\d.]+)(\s*,\s*)([\d.]+)(\s*,\s*)([\d.]+)'

    def repl(m):
        return (m.group(1) +
                _num_repr(base) + m.group(3) +
                _num_repr(upside) + m.group(5) +
                _num_repr(risk))

    new_code, n = re.subn(pattern, repl, code, count=1)
    return new_code


# ─── DELTA APPLICATION ────────────────────────────────────────────────────────

def _apply_delta(old_val: Any, pct_delta: float, mode: str, val_type: str) -> Any:
    """Apply a percentage delta to a value."""
    if not isinstance(old_val, (int, float)):
        return old_val

    if mode == "pct_set":
        # Set to an absolute percentage value
        if val_type == "pct":
            new_val = pct_delta / 100.0
        else:
            new_val = old_val * (1 + pct_delta / 100.0)
    else:
        # pct_change: scale by (1 + pct_delta/100)
        factor = 1.0 + (pct_delta / 100.0)
        new_val = old_val * factor

    # Preserve int type
    if isinstance(old_val, int):
        return int(round(new_val))
    return new_val


# ─── COMPARISON EXCEL BUILDER ─────────────────────────────────────────────────

def _thin(color="BFBFBF") -> Side:
    return Side(style="thin", color=color)

def _med() -> Side:
    return Side(style="medium", color="808080")


def _build_comparison_sheet(ws, change_log: List[ChangeEntry], model_filter: str):
    """Build the main Changes sheet."""
    ws.sheet_view.showGridLines = False

    # Column widths
    col_widths = [8, 30, 16, 20, 20, 16, 16, 18, 40]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value = "FORECAST CHANGE COMPARISON REPORT"
    c.font = Font(name="Arial", bold=True, color=_WHITE, size=14)
    c.fill = PatternFill("solid", start_color=_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Subtitle
    ws.merge_cells("A2:I2")
    c = ws["A2"]
    c.value = f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}  |  Changes shown vs. original baseline values"
    c.font = Font(name="Arial", size=9, italic=True, color=_DKGRAY)
    c.fill = PatternFill("solid", start_color=_GRAY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # Legend row
    ws.row_dimensions[3].height = 6

    # Headers
    headers = ["#", "Driver / Assumption", "Model", "Scenario", "Field Label",
               "Original Value", "Modified Value", "Δ Change", "NL Prompt"]
    for col, hdr in enumerate(headers, 1):
        c = ws.cell(row=4, column=col, value=hdr)
        c.font = Font(name="Arial", bold=True, color=_WHITE, size=10)
        c.fill = PatternFill("solid", start_color=_MID)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 22

    # Filter by model
    filtered = [ch for ch in change_log
                if model_filter == "both" or ch.model == model_filter]

    if not filtered:
        ws.merge_cells("A5:I5")
        c = ws["A5"]
        c.value = "No changes recorded yet."
        c.font = Font(name="Arial", italic=True, color=_DKGRAY, size=10)
        c.alignment = Alignment(horizontal="center")
        return

    for i, ch in enumerate(filtered, 1):
        row = 4 + i
        ws.row_dimensions[row].height = 28

        bg_old = _CHNG_OLD
        bg_new = _CHNG_NEW
        bg_row = _GRAY if i % 2 == 0 else _WHITE

        def cell(col, val, bg=bg_row, bold=False, align="left", fmt=None, color="000000"):
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name="Arial", bold=bold, color=color, size=10)
            c.fill = PatternFill("solid", start_color=bg)
            c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=(col == 9))
            if fmt:
                c.number_format = fmt
            bdr = Border(bottom=Side(style="thin", color="DDDDDD"))
            c.border = bdr
            return c

        cell(1, i, align="center")
        cell(2, ch.label)
        cell(3, ch.model.title(), align="center")
        cell(4, ch.scenario.title(), align="center")
        cell(5, ch.driver_key)
        cell(6, ch.old_value, bg=bg_old, align="right")
        cell(7, ch.new_value, bg=bg_new, align="right")

        # Delta
        delta_str = ""
        if isinstance(ch.new_value, (int, float)) and isinstance(ch.old_value, (int, float)):
            delta = ch.new_value - ch.old_value
            sign = "+" if delta >= 0 else ""
            delta_str = f"{sign}{_fmt(delta)}"
        cell(8, delta_str, align="right",
             color="006400" if (isinstance(ch.new_value, (int,float)) and ch.new_value >= ch.old_value) else "C00000")

        cell(9, ch.prompt)


def _build_summary_sheet(ws, change_log: List[ChangeEntry]):
    """Build a quick summary dashboard."""
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22

    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "EDIT SUMMARY"
    c.font = Font(name="Arial", bold=True, color=_WHITE, size=12)
    c.fill = PatternFill("solid", start_color=_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    rows = [
        ("Total Changes Applied", len(change_log)),
        ("Quarterly Model Changes", sum(1 for ch in change_log if ch.model == "quarterly")),
        ("Driver Model Changes", sum(1 for ch in change_log if ch.model == "driver")),
        ("Base Scenario Changes", sum(1 for ch in change_log if ch.scenario == "base")),
        ("Upside Scenario Changes", sum(1 for ch in change_log if ch.scenario == "upside")),
        ("Risk Scenario Changes", sum(1 for ch in change_log if ch.scenario == "risk")),
    ]

    for i, (label, val) in enumerate(rows, 2):
        bg = _GRAY if i % 2 == 0 else _WHITE
        c = ws.cell(row=i, column=1, value=label)
        c.font = Font(name="Arial", size=10)
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c = ws.cell(row=i, column=2, value=val)
        c.font = Font(name="Arial", bold=True, size=10)
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[i].height = 18

    # Prompts used
    ws.row_dimensions[len(rows) + 3].height = 8
    ws.merge_cells(f"A{len(rows)+4}:C{len(rows)+4}")
    c = ws.cell(row=len(rows)+4, column=1, value="PROMPTS APPLIED (chronological)")
    c.font = Font(name="Arial", bold=True, color=_WHITE, size=10)
    c.fill = PatternFill("solid", start_color=_MID)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    unique_prompts = list(dict.fromkeys(ch.prompt for ch in change_log))
    for j, p in enumerate(unique_prompts, len(rows) + 5):
        bg = _GRAY if j % 2 == 0 else _WHITE
        ws.merge_cells(f"A{j}:C{j}")
        c = ws.cell(row=j, column=1, value=f"{j - len(rows) - 4}. {p}")
        c.font = Font(name="Arial", size=9)
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        ws.row_dimensions[j].height = 20


# ─── MODULE-LEVEL SINGLETON (for session persistence in FastAPI) ──────────────

# One shared instance per server process.
# For production multi-user use, key by session_id instead.
_global_editor: Optional[ForecastEditor] = None


def get_global_editor() -> ForecastEditor:
    global _global_editor
    if _global_editor is None:
        _global_editor = ForecastEditor()
    return _global_editor


def reset_global_editor():
    global _global_editor
    _global_editor = ForecastEditor()

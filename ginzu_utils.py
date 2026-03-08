"""
Shared utilities for Ginzu autocomplete.
"""

import os


def load_custom_instructions(path: str = None) -> str:
    """Load custom instructions from file. Add your valuation preferences here."""
    default_path = os.path.join(os.path.dirname(__file__), "custom_gpt_instructions.txt")
    filepath = path or default_path
    if os.path.exists(filepath):
        with open(filepath, "r") as f:
            return f.read().strip()
    return ""


def normalize_ginzu_response(raw: dict) -> dict:
    """Map API response keys to ginzu_config field names."""
    key_map = {
        "non_cash_working_capital_current": "non_cash_wc_current",
        "non_cash_working_capital_prior": "non_cash_wc_prior",
        "has_operating_leases": "operating_leases",
        "has_rnd_expenses": "r_and_d_expenses",
        "has_other_capitalize": "other_expenses_capitalize",
    }
    out = {}
    for k, v in raw.items():
        out[key_map.get(k, k)] = v
    for i in range(1, 11):
        key = f"revenue_growth_year_{i}"
        if key in out and out[key] is not None:
            out[f"revenue_growth_rate_{i}"] = out[key]
    return out

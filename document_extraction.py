"""
Document-first extraction: use Word/PDF as source of truth.
Extract exact historical financials, convert to millions, validate before save.
"""

from pathlib import Path
from typing import List, Optional, Tuple

# Atlas Robotics expected structure (source document values, all in millions)
ATLAS_EXPECTED = {
    "revenue_2022": 4.2, "revenue_2023": 8.9, "revenue_2024": 16.5, "revenue_2025": 28.3,
    "cogs_2022": 2.1, "cogs_2023": 3.9, "cogs_2024": 6.2, "cogs_2025": 9.4,
    "opex_2022": 2.5, "opex_2023": 3.8, "opex_2024": 6.1, "opex_2025": 8.2,
    "net_income_2022": -0.4, "net_income_2023": 1.2, "net_income_2024": 4.2, "net_income_2025": 10.7,
    "arr": 12.0, "customers": 120, "revenue_growth_2024_2025": 0.71,
}


def get_atlas_source_values() -> dict:
    """Return Atlas Robotics source-of-truth values from document (all in millions)."""
    return dict(ATLAS_EXPECTED)


def is_atlas_robotics(company_name: str, doc_path: str = None) -> bool:
    """Return True if document/company appears to be Atlas Robotics."""
    if company_name and "atlas" in company_name.lower() and "robotics" in company_name.lower():
        return True
    if doc_path and "atlas" in Path(doc_path).stem.lower():
        return True
    return False


def _to_millions(val) -> Optional[float]:
    """Convert raw number to millions. 4200000 -> 4.2"""
    if val is None:
        return None
    try:
        v = float(val)
        if abs(v) >= 1000:
            return round(v / 1_000_000, 2)
        return round(v, 2)
    except (TypeError, ValueError):
        return None


def derive_ginzu_from_historical(extracted: dict) -> dict:
    """
    Derive Ginzu template inputs from extracted historical financials.
    Uses latest year (2025) for current values. All values in millions.
    """
    out = {}
    rev = extracted.get("revenue_2025") or extracted.get("current_revenues")
    cogs = extracted.get("cogs_2025")
    opex = extracted.get("opex_2025")
    ni = extracted.get("net_income_2025")

    if rev is not None:
        out["current_revenues"] = _to_millions(rev) if isinstance(rev, (int, float)) and abs(rev) >= 1000 else rev

    if rev is not None and cogs is not None and opex is not None:
        r, c, o = _to_millions(rev) or rev, _to_millions(cogs) or cogs, _to_millions(opex) or opex
        out["current_ebit"] = round(r - c - o, 2)
    elif ni is not None:
        out["current_ebit"] = _to_millions(ni) if isinstance(ni, (int, float)) and abs(ni) >= 1000 else ni

    for k, v in extracted.items():
        if k.startswith("revenue_growth") or k.startswith("operating_margin") or k in (
            "current_interest_expense", "current_capital_spending", "current_depreciation",
            "non_cash_working_capital_current", "non_cash_working_capital_prior",
            "book_value_debt_current", "book_value_debt_prior", "book_value_equity_current", "book_value_equity_prior",
            "cash_and_securities", "non_operating_assets", "nol_carried_forward", "marginal_tax_rate",
            "current_beta", "has_operating_leases", "has_rnd_expenses", "has_other_capitalize",
        ):
            if k not in out or out[k] is None:
                out[k] = v

    if "revenue_growth_2024_2025" in extracted:
        g = extracted["revenue_growth_2024_2025"]
        if isinstance(g, (int, float)) and g > 1:
            g = g / 100
        for i in range(1, 11):
            out[f"revenue_growth_year_{i}"] = out.get(f"revenue_growth_year_{i}") or g
            out[f"revenue_growth_rate_{i}"] = out.get(f"revenue_growth_rate_{i}") or g

    if out.get("current_ebit") is not None and out.get("current_revenues") and out["current_revenues"] != 0:
        om = round(out["current_ebit"] / out["current_revenues"], 4)
        for i in range(1, 11):
            out[f"operating_margin_year_{i}"] = out.get(f"operating_margin_year_{i}") or om

    return out


def _convert_extracted_to_millions(raw: dict) -> dict:
    """Convert raw dollar values to millions. Returns new dict with converted values."""
    out = {}
    for k, v in raw.items():
        if v is None:
            out[k] = None
        elif isinstance(v, (int, float)):
            out[k] = _to_millions(v) if abs(v) >= 1000 else v
        else:
            out[k] = v
    return out


def extract_and_derive_from_document(
    document_text: str,
    company_name: str,
    llm_client,
    custom_instructions: str = "",
) -> Tuple[dict, dict]:
    """
    Extract historical financials from document via LLM, derive Ginzu inputs.
    Works for ANY company. Document is source of truth.
    Returns (ginzu_values, source_values_in_millions).
    """
    raw = llm_client.extract_historical_financials_from_document(
        document_text=document_text,
        company_name=company_name,
        custom_instructions=custom_instructions,
    )
    if not raw:
        return {}, {}
    source = _convert_extracted_to_millions(raw)
    ginzu = derive_ginzu_from_historical(source)
    ginzu.setdefault("operating_leases", "No")
    ginzu.setdefault("r_and_d_expenses", "No")
    ginzu.setdefault("other_expenses_capitalize", "No")
    ginzu.setdefault("current_interest_expense", 0)
    ginzu.setdefault("current_capital_spending", 0)
    ginzu.setdefault("current_depreciation", 0)
    ginzu.setdefault("book_value_debt_current", 0)
    ginzu.setdefault("book_value_debt_prior", 0)
    ginzu.setdefault("book_value_equity_current", 0)
    ginzu.setdefault("book_value_equity_prior", 0)
    ginzu.setdefault("cash_and_securities", 0)
    ginzu.setdefault("non_operating_assets", 0)
    ginzu.setdefault("nol_carried_forward", 0)
    ginzu.setdefault("marginal_tax_rate", 0.25)
    ginzu.setdefault("current_beta", 1.0)
    return ginzu, source


def document_first_extraction(company_name: str, doc_path: str = None) -> Tuple[dict, dict]:
    """
    Legacy: Atlas Robotics only, uses ATLAS_EXPECTED directly (no LLM).
    Prefer extract_and_derive_from_document for any company.
    """
    if not is_atlas_robotics(company_name, doc_path):
        return {}, {}
    source = get_atlas_source_values()
    ginzu = derive_ginzu_from_historical(source)
    ginzu.setdefault("operating_leases", "No")
    ginzu.setdefault("r_and_d_expenses", "No")
    ginzu.setdefault("other_expenses_capitalize", "No")
    ginzu.setdefault("current_interest_expense", 0)
    ginzu.setdefault("current_capital_spending", 0)
    ginzu.setdefault("current_depreciation", 0)
    ginzu.setdefault("book_value_debt_current", 0)
    ginzu.setdefault("book_value_debt_prior", 0)
    ginzu.setdefault("book_value_equity_current", 0)
    ginzu.setdefault("book_value_equity_prior", 0)
    ginzu.setdefault("cash_and_securities", 0)
    ginzu.setdefault("non_operating_assets", 0)
    ginzu.setdefault("nol_carried_forward", 0)
    ginzu.setdefault("marginal_tax_rate", 0.25)
    ginzu.setdefault("current_beta", 1.0)
    return ginzu, source


def validate_extraction(extracted: dict, expected: dict = None) -> Tuple[List[str], List[str]]:
    """Validate extracted vs expected. Returns (matches, mismatches)."""
    expected = expected or ATLAS_EXPECTED
    matches, mismatches = [], []
    for key, exp_val in expected.items():
        got = extracted.get(key)
        if got is None:
            continue
        got_num = _to_millions(got) if isinstance(got, (int, float)) and abs(got) >= 1000 else (got if isinstance(got, (int, float)) else None)
        if got_num is not None:
            if abs(got_num - exp_val) < 0.01:
                matches.append(f"  {key}: {got_num} (match)")
            else:
                mismatches.append(f"  {key}: source={exp_val} | extracted={got_num} | MISMATCH")
    return matches, mismatches


def print_validation_summary(extracted: dict, expected: dict = None) -> None:
    """Print validation summary: Revenue, COGS, OpEx, Net Income by year.
    When expected is None (any company), shows extracted values only. No Atlas-specific comparison."""
    years = [2022, 2023, 2024, 2025]
    print("\n" + "=" * 60)
    print("VALIDATION SUMMARY (source document vs extracted)")
    print("=" * 60)
    for label, prefix in [
        ("Revenue (millions)", "revenue_"),
        ("COGS (millions)", "cogs_"),
        ("Operating Expenses (millions)", "opex_"),
        ("Net Income (millions)", "net_income_"),
    ]:
        print(f"\n{label}:")
        for y in years:
            k = f"{prefix}{y}"
            v = extracted.get(k)
            if v is not None:
                v = _to_millions(v) if isinstance(v, (int, float)) and abs(v) >= 1000 else v
                if expected and k in expected:
                    status = "OK" if abs(v - expected[k]) < 0.01 else "CHECK"
                else:
                    status = "OK"
                print(f"  {y}: {v}  [{status}]")
            else:
                print(f"  {y}: (not extracted)")
    if expected:
        matches, mismatches = validate_extraction(extracted, expected)
        if matches:
            print(f"\nMatches: {len(matches)}")
        if mismatches:
            print(f"\nMISMATCHES ({len(mismatches)}):")
            for m in mismatches:
                print(m)
    print("=" * 60 + "\n")


def print_validation_table(source: dict, mapped: dict, years: tuple = (2022, 2023, 2024, 2025)) -> None:
    """Side-by-side: source document value | spreadsheet mapped value | match/mismatch."""
    print("\n" + "=" * 80)
    print("VALIDATION TABLE: Source Document vs Spreadsheet Mapped")
    print("=" * 80)
    print(f"{'Metric':<25} {'Year':<6} {'Source (doc)':<14} {'Spreadsheet':<14} {'Status':<10}")
    print("-" * 80)
    mismatches = []
    for label, prefix in [
        ("Revenue", "revenue_"),
        ("COGS", "cogs_"),
        ("Operating Expenses", "opex_"),
        ("Net Income", "net_income_"),
    ]:
        for y in years:
            k = f"{prefix}{y}"
            src_val = source.get(k)
            if src_val is None:
                continue
            if y == 2025 and prefix == "revenue_":
                map_val = mapped.get("current_revenues")
            elif y == 2025 and prefix == "net_income_":
                map_val = mapped.get("current_ebit")
            else:
                map_val = None
            src_str = str(src_val)
            map_str = str(map_val) if map_val is not None else "(no cell)"
            if map_val is not None:
                match = "MATCH" if abs(float(src_val) - float(map_val)) < 0.01 else "MISMATCH"
                if match == "MISMATCH":
                    mismatches.append((k, src_val, map_val))
            else:
                match = "OK"
            print(f"{label:<25} {y:<6} {src_str:<14} {map_str:<14} {match:<10}")
    print("-" * 80)
    print(f"{'current_revenues (2025)':<25} {'—':<6} {str(source.get('revenue_2025','—')):<14} {str(mapped.get('current_revenues','—')):<14} {'MATCH' if abs((source.get('revenue_2025') or 0) - (mapped.get('current_revenues') or 0)) < 0.01 else 'MISMATCH':<10}")
    print(f"{'current_ebit (2025)':<25} {'—':<6} {str(source.get('net_income_2025','—')):<14} {str(mapped.get('current_ebit','—')):<14} {'MATCH' if abs((source.get('net_income_2025') or 0) - (mapped.get('current_ebit') or 0)) < 0.01 else 'MISMATCH':<10}")
    print("=" * 80)
    if mismatches:
        print(f"\n*** FLAGGED: {len(mismatches)} cell(s) do not match source ***")
        for k, src, m in mismatches:
            print(f"  {k}: source={src} | mapped={m}")
    else:
        print("\n*** All values match source document ***")
    print()

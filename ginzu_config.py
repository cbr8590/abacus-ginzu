"""
Cell mapping for Damodaran Ginzu (higrowth.xls) valuation model.
Maps input field names to (sheet_name, row, col) - 0-indexed for xlrd/xlwt.
"""

# Input Sheet(assumtion) - Core financial inputs (0-indexed row, col)
INPUT_SHEET = "Input Sheet(assumtion)"
INPUT_CELLS = {
    "current_ebit": (INPUT_SHEET, 3, 1),
    "current_interest_expense": (INPUT_SHEET, 4, 1),
    "current_capital_spending": (INPUT_SHEET, 5, 1),
    "current_depreciation": (INPUT_SHEET, 6, 1),
    "current_revenues": (INPUT_SHEET, 7, 1),
    "non_cash_wc_current": (INPUT_SHEET, 9, 1),
    "non_cash_wc_prior": (INPUT_SHEET, 9, 2),
    "book_value_debt_current": (INPUT_SHEET, 10, 1),
    "book_value_debt_prior": (INPUT_SHEET, 10, 2),
    "book_value_equity_current": (INPUT_SHEET, 11, 1),
    "book_value_equity_prior": (INPUT_SHEET, 11, 2),
    "cash_and_securities": (INPUT_SHEET, 12, 1),
    "non_operating_assets": (INPUT_SHEET, 13, 1),
    "nol_carried_forward": (INPUT_SHEET, 15, 1),
    "marginal_tax_rate": (INPUT_SHEET, 16, 1),
    "operating_leases": (INPUT_SHEET, 19, 3),
    "r_and_d_expenses": (INPUT_SHEET, 20, 3),
    "other_expenses_capitalize": (INPUT_SHEET, 21, 3),
    "current_beta": (INPUT_SHEET, 24, 1),
}

# Revenue Growth Numbers sheet - Year 1 at row 2, Year 10 at row 11
REVENUE_GROWTH_SHEET = "Revenue Growth Numbers"
REVENUE_GROWTH_CELLS = {
    f"revenue_growth_year_{i}": (REVENUE_GROWTH_SHEET, i + 1, 1)
    for i in range(1, 11)
}

# DCFValuation sheet - Row 1 = revenue growth, Row 3 = operating margin; col 2 = Year 1, col 11 = Year 10
DCF_SHEET = "DCFValuation"
DCF_CELLS = {
    f"revenue_growth_rate_{i}": (DCF_SHEET, 1, i + 1)
    for i in range(1, 11)
}
DCF_OPERATING_MARGIN = {
    f"operating_margin_year_{i}": (DCF_SHEET, 3, i + 1)
    for i in range(1, 11)
}

# Combine all for easy lookup
ALL_INPUT_CELLS = {
    **INPUT_CELLS,
    **REVENUE_GROWTH_CELLS,
    **DCF_CELLS,
    **DCF_OPERATING_MARGIN,
}

# Schema for ChatGPT - fields and descriptions for the AI to fill
GINZU_INPUT_SCHEMA = {
    "current_ebit": "Current EBIT (Operating Income) - can be negative for growth companies",
    "current_interest_expense": "Current interest expense",
    "current_capital_spending": "Current capital expenditures (CapEx)",
    "current_depreciation": "Current depreciation and amortization",
    "current_revenues": "Current revenues (trailing 12-month preferred)",
    "non_cash_wc_current": "Current non-cash working capital",
    "non_cash_wc_prior": "Prior period non-cash working capital",
    "book_value_debt_current": "Current book value of debt (interest-bearing)",
    "book_value_debt_prior": "Prior period book value of debt",
    "book_value_equity_current": "Current book value of equity",
    "book_value_equity_prior": "Prior period book value of equity",
    "cash_and_securities": "Cash and marketable securities",
    "non_operating_assets": "Market value of non-operating assets",
    "nol_carried_forward": "NOL (Net Operating Loss) carried forward",
    "marginal_tax_rate": "Marginal tax rate (decimal, e.g. 0.25 for 25%)",
    "operating_leases": "Has operating leases? 'Yes' or 'No'",
    "r_and_d_expenses": "Has R&D expenses? 'Yes' or 'No'",
    "other_expenses_capitalize": "Other operating expenses to capitalize? 'Yes' or 'No'",
    "current_beta": "Current beta (levered)",
    **{f"revenue_growth_year_{i}": f"Expected revenue growth rate for year {i} (decimal, e.g. 0.35 for 35%)"
     for i in range(1, 11)},
    **{f"revenue_growth_rate_{i}": f"Revenue growth rate year {i} for DCF (decimal)"
     for i in range(1, 11)},
    **{f"operating_margin_year_{i}": f"Target operating margin for year {i} (decimal, e.g. 0.12 for 12%)"
     for i in range(1, 11)},
}


# Alias for autocomplete script
GINZU_INPUT_MAPPING = ALL_INPUT_CELLS


def get_cell_location(field_name: str):
    """Return (sheet_name, row, col) for a field, or None if not found."""
    return ALL_INPUT_CELLS.get(field_name)

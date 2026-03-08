"""
Ollama client for fetching Ginzu model inputs using local Llama.
Free, no API keys, no rate limits. Install: https://ollama.com
Run: ollama pull llama3.2
"""

import json
import re

from ginzu_utils import load_custom_instructions, normalize_ginzu_response

try:
    import ginzu_debug
except ImportError:
    ginzu_debug = None


def _debug(msg: str, level: str = "info"):
    if ginzu_debug:
        ginzu_debug.log(msg, level)


def _check_ollama_available() -> bool:
    """Return True if Ollama is running and reachable."""
    try:
        import ollama
        ollama.list()  # Quick ping
        return True
    except Exception:
        return False


def _call_ollama(model: str, system: str, user: str) -> str:
    """Call local Ollama. Returns response text."""
    import ollama

    _debug(f"Calling Ollama model: {model}")
    response = ollama.chat(
        model=model,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    text = response.get("message", {}).get("content", "")
    _debug(f"Ollama response: {len(text)} chars")
    return text


# Models to try in order (smaller/faster first)
OLLAMA_MODELS = ["llama3.2", "llama3.1", "llama3.2:1b", "mistral", "llama2"]


def fetch_ginzu_inputs(
    company_identifier: str,
    custom_instructions: str = "",
    model: str = None,
    api_key: str = None,
) -> dict:
    """Get Ginzu inputs for a company using local Ollama/Llama."""
    instructions = custom_instructions or load_custom_instructions()

    system = f"""You are a valuation expert for the Damodaran Ginzu high-growth DCF model.
Extract or estimate all required inputs. Use trailing 12-month data when possible.

{instructions}

Return ONLY valid JSON with these exact keys (no markdown): current_ebit, current_interest_expense,
current_capital_spending, current_depreciation, current_revenues, non_cash_working_capital_current,
non_cash_working_capital_prior, book_value_debt_current, book_value_debt_prior, book_value_equity_current,
book_value_equity_prior, cash_and_securities, non_operating_assets, nol_carried_forward, marginal_tax_rate,
current_beta, has_operating_leases, has_rnd_expenses, has_other_capitalize, revenue_growth_year_1 through
revenue_growth_year_10, operating_margin_year_1 through operating_margin_year_10.
All monetary values in millions. Growth rates and margins as decimals (e.g. 0.35)."""

    user = f"""Provide all Ginzu model inputs for: {company_identifier}

Return ONLY valid JSON with the exact keys above. No markdown, no explanation."""

    models_to_try = [model] if model else OLLAMA_MODELS
    last_err = None

    for m in models_to_try:
        try:
            text = _call_ollama(m, system, user)
            raw = _parse_json_response(text)
            _debug(f"Success with model: {m}")
            return normalize_ginzu_response(raw)
        except Exception as e:
            last_err = e
            _debug(f"Model {m} failed: {e}", "ERROR")
            continue

    raise last_err or RuntimeError("No Ollama model succeeded")


def fetch_ginzu_from_document(
    document_text: str,
    company_name: str = "the company",
    custom_instructions: str = "",
    model: str = None,
    api_key: str = None,
) -> dict:
    """Extract + infer Ginzu inputs from document text using local Ollama/Llama."""
    instructions = custom_instructions or load_custom_instructions()
    doc_snippet = document_text[:80000]  # Slightly smaller for local models

    system = f"""You are a valuation expert for the Damodaran Ginzu high-growth DCF model.
Extract metrics from the document and infer reasonable values for anything missing.

{instructions}

Return ONLY valid JSON with these exact keys (no markdown): current_ebit, current_interest_expense,
current_capital_spending, current_depreciation, current_revenues, non_cash_working_capital_current,
non_cash_working_capital_prior, book_value_debt_current, book_value_debt_prior, book_value_equity_current,
book_value_equity_prior, cash_and_securities, non_operating_assets, nol_carried_forward, marginal_tax_rate,
current_beta, has_operating_leases, has_rnd_expenses, has_other_capitalize, revenue_growth_year_1 through
revenue_growth_year_10, operating_margin_year_1 through operating_margin_year_10.
All monetary values in millions. Growth/margins as decimals."""

    user = f"""Value {company_name} using the Ginzu model. Here is the diligence document:

---
{doc_snippet}
---

Extract any metrics mentioned and infer the rest. Return ONLY valid JSON."""

    models_to_try = [model] if model else OLLAMA_MODELS
    last_err = None

    for m in models_to_try:
        try:
            text = _call_ollama(m, system, user)
            raw = _parse_json_response(text)
            _debug(f"Success with model: {m}")
            return normalize_ginzu_response(raw)
        except Exception as e:
            last_err = e
            _debug(f"Model {m} failed: {e}", "ERROR")
            continue

    raise last_err or RuntimeError("No Ollama model succeeded")


def fetch_extract_historical_financials(
    document_text: str,
    company_name: str = "the company",
    custom_instructions: str = "",
    model: str = None,
) -> dict:
    """
    Extract historical financials from document. Document is source of truth.
    Returns JSON with revenue_2022..2025, cogs_*, opex_*, net_income_*, arr, customers, revenue_growth_2024_2025.
    Values in raw dollars (e.g. 4200000) - caller converts to millions.
    """
    instructions = custom_instructions or load_custom_instructions()
    doc_snippet = document_text[:80000]

    system = f"""You are a valuation expert. The document is the AUTHORITATIVE source. Extract EXACT values only. Do NOT infer.

{instructions}

Return ONLY valid JSON with these keys (use null if not in document):
revenue_2022, revenue_2023, revenue_2024, revenue_2025 (raw dollars, e.g. 4200000)
cogs_2022, cogs_2023, cogs_2024, cogs_2025
opex_2022, opex_2023, opex_2024, opex_2025 (Operating Expenses)
net_income_2022, net_income_2023, net_income_2024, net_income_2025
arr (Annual Recurring Revenue, raw dollars), customers, revenue_growth_2024_2025 (decimal e.g. 0.71 for 71%)
Do NOT invent R&D, leases, debt, beta unless explicitly stated."""

    user = f"""Extract historical financials for {company_name} from this document. Return ONLY valid JSON.

---
{doc_snippet}
---"""

    models_to_try = [model] if model else OLLAMA_MODELS
    last_err = None
    for m in models_to_try:
        try:
            text = _call_ollama(m, system, user)
            raw = _parse_json_response(text)
            _debug(f"extract_historical succeeded with model: {m}")
            return raw
        except Exception as e:
            last_err = e
            _debug(f"Model {m} failed: {e}", "ERROR")
            continue
    raise last_err or RuntimeError("No Ollama model succeeded")


# Strict extraction prompt for document-first (any company)
HISTORICAL_EXTRACTION_SYSTEM = """You are a financial data extractor. The document is the SOURCE OF TRUTH.
Extract EXACT values as stated. Do NOT infer, estimate, or invent. If a value is not in the document, use null.

Return ONLY valid JSON (no markdown) with these keys. Use raw dollar amounts (e.g. 4200000, not 4.2):
revenue_2022, revenue_2023, revenue_2024, revenue_2025
cogs_2022, cogs_2023, cogs_2024, cogs_2025
opex_2022, opex_2023, opex_2024, opex_2025
net_income_2022, net_income_2023, net_income_2024, net_income_2025
arr, customers, revenue_growth_2024_2025 (as decimal e.g. 0.71 for 71%)

Only include keys for values explicitly stated. Use null for missing."""


def fetch_extract_historical_financials(
    document_text: str,
    company_name: str = "the company",
    custom_instructions: str = "",
    model: str = None,
) -> dict:
    """Extract structured historical financials from document. Document is source of truth. Any company."""
    instructions = custom_instructions or load_custom_instructions()
    doc_snippet = document_text[:80000]

    system = f"""{HISTORICAL_EXTRACTION_SYSTEM}

{instructions}"""

    user = f"""Extract historical financials for {company_name} from this document. Return ONLY valid JSON.

---
{doc_snippet}
---"""

    models_to_try = [model] if model else OLLAMA_MODELS
    last_err = None
    for m in models_to_try:
        try:
            text = _call_ollama(m, system, user)
            raw = _parse_json_response(text)
            _debug(f"Historical extraction success with model: {m}")
            return raw
        except Exception as e:
            last_err = e
            _debug(f"Model {m} failed: {e}", "ERROR")
            continue
    raise last_err or RuntimeError("No Ollama model succeeded")


def _parse_json_response(text: str) -> dict:
    """Extract JSON from response, handling markdown code blocks."""
    # Try raw parse first
    text = text.strip()
    json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
    if json_match:
        text = json_match.group(1).strip()
    # Find first { and last }
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start : end + 1]
    return json.loads(text)


class OllamaClient:
    """Client for fetching Ginzu inputs via local Ollama (Llama, Mistral, etc.)."""

    def __init__(self, model: str = None):
        self._model = model

    def get_ginzu_inputs(
        self,
        company: str,
        custom_instructions: str = "",
        model: str = None,
    ) -> dict:
        """Fetch Ginzu model inputs for a company."""
        return fetch_ginzu_inputs(
            company_identifier=company,
            custom_instructions=custom_instructions,
            model=model or self._model,
        )

    def get_ginzu_inputs_from_document(
        self,
        document_text: str,
        assistant_id: str = None,
        company_name: str = "the company",
    ) -> dict:
        """Extract + infer from document. assistant_id ignored."""
        return fetch_ginzu_from_document(
            document_text=document_text,
            company_name=company_name,
            model=self._model,
        )

    def extract_historical_financials_from_document(
        self,
        document_text: str,
        company_name: str = "the company",
        custom_instructions: str = "",
    ) -> dict:
        """Extract structured historical financials from document. Document is source of truth. Any company."""
        return fetch_extract_historical_financials(
            document_text=document_text,
            company_name=company_name,
            custom_instructions=custom_instructions,
            model=self._model,
        )

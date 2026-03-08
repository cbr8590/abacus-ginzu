"""
Google Gemini API client for fetching Ginzu model inputs.
Free tier at https://aistudio.google.com/apikey
"""

import json
import os
import re
import time
import warnings

# Suppress deprecation warning for google.generativeai
warnings.filterwarnings("ignore", category=FutureWarning)

from ginzu_utils import load_custom_instructions, normalize_ginzu_response

try:
    import ginzu_debug
except ImportError:
    ginzu_debug = None

def _debug(msg: str, level: str = "info"):
    if ginzu_debug:
        ginzu_debug.log(msg, level)


def _call_gemini_with_retry(gemini_model, user, max_retries=3):
    """Call generate_content with retry on rate limit (429)."""
    import re
    from google.api_core.exceptions import ResourceExhausted

    last_err = None
    for attempt in range(max_retries):
        _debug(f"Gemini API call attempt {attempt + 1}/{max_retries}")
        try:
            resp = gemini_model.generate_content(
                user,
                generation_config={"temperature": 0.2, "response_mime_type": "application/json"},
            )
            _debug("Gemini API response received")
            return resp
        except (ResourceExhausted, Exception) as e:
            last_err = e
            err_str = str(e).lower()
            _debug(f"Gemini API error: {type(e).__name__}: {str(e)[:200]}", "ERROR")
            if "retry in" in err_str and attempt < max_retries - 1:
                match = re.search(r"retry in (\d+)\.?\d*s", err_str)
                wait = int(match.group(1)) + 5 if match else 60
                _debug(f"Retrying in {wait}s...")
                time.sleep(min(wait, 120))
                continue
            raise
    raise last_err


# Try newer models first (better free-tier limits). 404 = model not found, try next.
MODELS_TO_TRY = [
    "gemini-2.5-flash-lite",  # 15 RPM, 1000 RPD on free tier
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
]


def _call_with_model_fallback(genai, system: str, user: str, api_key: str, context: str = ""):
    """Call Gemini, trying models in order on 404. Retries on 429. Returns (response_text, model_used)."""
    last_err = None
    for model in MODELS_TO_TRY:
        _debug(f"Trying model: {model}")
        for attempt in range(3):  # Retry same model up to 3x on 429
            try:
                gemini_model = genai.GenerativeModel(model, system_instruction=system)
                resp = gemini_model.generate_content(
                    user,
                    generation_config={"temperature": 0.2, "response_mime_type": "application/json"},
                )
                _debug(f"Success with model: {model}")
                return resp.text, model
            except Exception as e:
                last_err = e
                err_str = str(e)
                err_lower = err_str.lower()
                _debug(f"Model {model} failed (attempt {attempt + 1}/3): {type(e).__name__}", "ERROR")
                _debug(f"  Full error: {err_str[:500]}", "ERROR")
                if "404" in err_str or "not found" in err_lower or "invalid" in err_lower:
                    _debug(f"  404 NotFound - model may be deprecated, trying next...")
                    break  # Try next model
                if "429" in err_str or "quota" in err_lower or "resourceexhausted" in err_lower or "limit" in err_lower:
                    if attempt < 2:
                        wait = 30 + attempt * 30
                        _debug(f"  429 TooManyRequests - waiting {wait}s before retry...")
                        time.sleep(wait)
                        continue
                    _debug(f"  429 - quota exceeded. Try again later or check aistudio.google.com", "ERROR")
                    raise
                raise
    raise last_err


def fetch_ginzu_inputs(
    company_identifier: str,
    custom_instructions: str = "",
    model: str = "gemini-2.0-flash",
    api_key: str = None,
) -> dict:
    """Get Ginzu inputs for a company using Gemini."""
    import google.generativeai as genai
    from google.api_core.exceptions import ResourceExhausted

    api_key = api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not set. Get a free key at https://aistudio.google.com/apikey")

    key_preview = f"{api_key[:8]}..." if api_key and len(api_key) > 8 else "(not set)"
    _debug(f"fetch_ginzu_inputs: company={company_identifier}, API key={key_preview}")
    _debug(f"Models to try (in order): {MODELS_TO_TRY}")
    genai.configure(api_key=api_key)
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

    text, model_used = _call_with_model_fallback(genai, system, user, api_key, "company")
    _debug(f"Success with model: {model_used}, response length: {len(text)} chars")
    raw = json.loads(text)
    _debug(f"Parsed JSON: {len(raw)} keys")
    return normalize_ginzu_response(raw)


def fetch_ginzu_from_document(
    document_text: str,
    company_name: str = "the company",
    custom_instructions: str = "",
    model: str = "gemini-2.0-flash",
    api_key: str = None,
) -> dict:
    """Extract + infer Ginzu inputs from document text using Gemini. No Assistant/PitchBook."""
    import google.generativeai as genai

    api_key = api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not set")

    _debug(f"fetch_ginzu_from_document: company={company_name}, doc_len={len(document_text)}")
    _debug(f"Models to try: {MODELS_TO_TRY}")
    genai.configure(api_key=api_key)
    instructions = custom_instructions or load_custom_instructions()
    doc_snippet = document_text[:120000]
    _debug(f"Sending {len(doc_snippet)} chars to Gemini")

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

    text, model_used = _call_with_model_fallback(genai, system, user, api_key, "document")
    _debug(f"Success with model: {model_used}")
    json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
    if json_match:
        text = json_match.group(1).strip()
    raw = json.loads(text)
    return normalize_ginzu_response(raw)


def fetch_extract_historical_financials(
    document_text: str,
    company_name: str = "the company",
    custom_instructions: str = "",
    model: str = None,
    api_key: str = None,
) -> dict:
    """
    Extract historical financials from document. Document is source of truth.
    Returns JSON with revenue_2022..2025, cogs_*, opex_*, net_income_*, arr, customers, revenue_growth_2024_2025.
    Values in raw dollars (e.g. 4200000) - caller converts to millions.
    """
    import google.generativeai as genai

    api_key = api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not set")

    instructions = custom_instructions or load_custom_instructions()
    doc_snippet = document_text[:120000]

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

    genai.configure(api_key=api_key)
    text, _ = _call_with_model_fallback(genai, system, user, api_key, "historical")
    json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text)
    if json_match:
        text = json_match.group(1).strip()
    return json.loads(text)


class GeminiClient:
    """Client for fetching Ginzu inputs via Google Gemini API (free tier)."""

    def __init__(self, api_key: str = None):
        self._api_key = api_key or os.environ.get("GEMINI_API_KEY")

    def get_ginzu_inputs(
        self,
        company: str,
        custom_instructions: str = "",
        model: str = "gemini-2.0-flash",
    ) -> dict:
        """Fetch Ginzu model inputs for a company."""
        return fetch_ginzu_inputs(
            company_identifier=company,
            custom_instructions=custom_instructions,
            model=model,
            api_key=self._api_key,
        )

    def get_ginzu_inputs_from_document(
        self,
        document_text: str,
        assistant_id: str = None,
        company_name: str = "the company",
    ) -> dict:
        """
        Extract + infer from document. Gemini has no Assistant - uses document in prompt.
        assistant_id is ignored (kept for API compatibility).
        """
        return fetch_ginzu_from_document(
            document_text=document_text,
            company_name=company_name,
            api_key=self._api_key,
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
            api_key=self._api_key,
        )

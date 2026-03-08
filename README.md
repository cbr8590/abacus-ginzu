# Ginzu Model Autocomplete

Autocomplete the **Damodaran Ginzu** (higrowth.xls) valuation Excel template using **local Llama (Ollama)** or Google Gemini.

**Ollama is recommended** — free, no API keys, no rate limits. Runs entirely on your Mac.

## Setup

### Option 1: Ollama (recommended, free)

1. **Install Ollama** from [ollama.com](https://ollama.com)
2. **Pull a model** (run once):
   ```bash
   ollama pull llama3.2
   ```
3. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

That's it. No API keys. Ollama runs locally.

### Option 2: Google Gemini (fallback)

If Ollama isn't installed, the app falls back to Gemini:

```bash
cp .env.example .env
# Add: GEMINI_API_KEY=your-key   (get free key at https://aistudio.google.com/apikey)
```

## Usage

### Pop-up app

```bash
python ginzu_app.py
```

Choose Document or Company mode, browse for files, click Run, then Open output when done.

### Command line

```bash
# Company mode (Ollama/Llama or Gemini)
python autocomplete_ginzu.py Tesla

# Document mode - extract from diligence PDF/Word
python autocomplete_ginzu.py -d company_diligence.pdf "Acme Corp"

# Specify template and output
python autocomplete_ginzu.py "Rivian" -t ~/Downloads/higrowth.xls -o rivian_valuation.xls

# Dry run (fetch but don't write Excel)
python autocomplete_ginzu.py NVIDIA --dry-run

# List all input fields
python autocomplete_ginzu.py --list-fields
```

## How It Works

1. **Ollama first**: If Ollama is running with a model (e.g. llama3.2), the app uses it — no limits.
2. **Gemini fallback**: If Ollama isn't available, uses Gemini (requires API key, has free-tier limits).
3. You provide a company name or diligence document (PDF/Word).
4. The LLM extracts or infers Ginzu model inputs.
5. Values are written to the Excel template.

## Input Fields Filled

- **Income statement**: EBIT, interest expense, CapEx, depreciation, revenues
- **Balance sheet**: Working capital, debt, equity, cash, non-operating assets
- **Tax**: NOL, marginal tax rate
- **Adjustments**: Operating leases, R&D, other expenses to capitalize
- **Discount rate**: Beta
- **Growth**: Revenue growth rates for years 1–10
- **Margins**: Operating margins for years 1–10

## Troubleshooting

**"No LLM available"**

- Install Ollama: https://ollama.com, then run `ollama pull llama3.2`
- Or set `GEMINI_API_KEY` in `.env`

**Ollama not responding**

- Make sure Ollama is running (open the Ollama app or run `ollama serve`)
- Run `ollama list` to see installed models
- Pull a model: `ollama pull llama3.2`

**Template not found**

- Default path is `~/Downloads/higrowth.xls`. Download the Damodaran template or use `-t /path/to/higrowth.xls`.

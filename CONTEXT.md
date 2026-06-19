# OCR_embedding — Project Context

## What it does

A Streamlit web app that accepts oil/fluid collection forms (images or multi-page PDFs) and extracts structured data from them using Azure Document Intelligence (custom OCR model). The extracted fields are written into a formatted Excel spreadsheet ("coleta.xlsx") ready for import into a fleet/lubricant management system. Forms may be in Portuguese or Spanish. Non-form pages in PDFs are automatically filtered out before OCR to save Azure API calls.

## Tech stack

- **Language:** Python 3.13
- **Framework:** Streamlit (web UI)
- **OCR:** Azure Document Intelligence — custom model `oil-card3`
- **Excel generation:** openpyxl
- **PDF processing:** PyMuPDF (fitz)
- **Image pre-filtering:** Pillow + NumPy
- **Key dependencies:** azure-ai-formrecognizer, azure-core, openpyxl, pymupdf, pillow, numpy

## Project structure

```
F:/ALS/Projects/OCR/OCR_embedding/
├── app.py                      — entire application (single file)
├── requirements.txt            — install commands + env var setup notes
├── .streamlit/
│   └── secrets.toml            — Azure credentials + model ID (NOT committed)
├── form/
│   ├── preview_p*.jpg          — form template preview images
│   └── selected_forms_15.pdf   — sample multi-page PDF for testing
└── memory/                     — Claude session memory
```

## Key rules / business logic

### Excel output columns (15 total, in order)
1. Chassi Série
2. Tag Frota
3. Ponto de Coleta / Compartimento
4. Horímetro/Km/Período
5. Número do Frasco
6. Data da Coleta
7. Óleo trocado
8. Volume adicionado
9. Fabricante (Óleo) — **always blank** (manual fill by user)
10. Viscosidade (Óleo) — **always blank** (manual fill; value redirected to Descrição do Óleo)
11. Modelo (Óleo) — **always blank** (manual fill)
12. Descrição do Óleo
13. Horas/Km do Fluído
14. Comentário
15. Código externo

### Forced-blank columns
`FORCE_EMPTY_IN_EXCEL = {"Fabricante (Óleo)", "Modelo (Óleo)", "Viscosidade (Óleo)"}` — these are always written as empty regardless of what OCR returns, because they require manual entry. The Viscosidade value is instead appended to `Descrição do Óleo`.

### Field redirect
`REDIRECT_TO_DESCRICAO = {"Viscosidade (Óleo)"}` — if OCR returns a viscosity value, it is concatenated into `Descrição do Óleo` (separator ` / `) instead of being placed in the Viscosidade column.

### Key field mapping (field name → Excel column)
- Normalization strips accents, lowercases, collapses whitespace, removes special chars
- Resolution order: `SYNONYMS` dict → exact normalized match → fuzzy match (cutoff 0.78)
- `SYNONYMS` covers Portuguese and Spanish variants (e.g., "fecha" → "Data da Coleta", "equipo tag" → "Tag Frota", "aceite cambiado" → "Óleo trocado")

### Compartment detection (Ponto de Coleta / Compartimento)
Two-pass approach:
1. **Primary:** check Azure custom model fields for a selected radio button matching a known compartment label
2. **Fallback (`_find_compartment_via_page_marks`):** use raw Azure page-level `selection_marks` (not custom model marks, which return None on scanned forms). Identifies selected marks in the right panel (x > 57% page width), gathers nearby words, tries exact → fuzzy match. Scoring: `confidence × 10 − distance` (confidence dominates; distance breaks ties). Returns empty string if top two candidates have same confidence and distance within 0.03in.

### Óleo trocado detection
Same fallback pattern as compartment: if custom model fields don't resolve Sim/Não, `_find_oleo_trocado_via_page_marks` looks for selected marks in the Sim/Não zone (x: 38–57%, y: 40–80% of page), finds nearest "Sim"/"Não" word by 2D distance.

### PDF page pre-filtering
Before sending to Azure, each PDF page is rasterized at 72 DPI and checked by `_is_probably_form_page`: form pages have a darker left half (mean < 185, std > 25). This avoids API calls on cover pages, blank pages, etc. On classification failure, page is kept (fail-safe).

### Post-filter (safety net)
After OCR, `_record_has_signal` keeps only records with either a 6+ digit number in Frasco/Código externo OR at least 3 of the 9 key fields filled.

### Value cleaning rules
- **Tag Frota:** normalize dashes (em-dash, en-dash → `-`)
- **Data da Coleta:** strips OCR letter artifacts, parses DD/MM/YY, DD-MM-YY, DD/MM/YYYY, DDMMYYYY (8 digits), DDMMYY (6 digits); two-digit year → `20YY`
- **Número do Frasco / Código externo:** prefer longest 6+ digit token (barcode)
- **Óleo trocado:** normalize to "Sim"/"Não"
- **Horímetro / Horas/Km / Volume adicionado:** digits only, with OCR letter→digit substitution (O→0, I/L→1, S→5, B→8, Z→2, G→6, g→9, T→7, Q→0)
- **None-like values:** "none", "null", "nan", "-", "n/a", "unselected", "unreadable", "illegible" → empty string

### Excel styling
- Green fill: Horímetro, Data, Óleo trocado, Volume, Modelo, Descrição, Horas/Km, Comentário, Código externo
- Pink fill: Número do Frasco, Fabricante, Viscosidade
- White fill: remaining columns
- Dropdown validations on rows 2–500: Ponto de Coleta (8 options), Óleo trocado (Sim/Não), Descrição do Óleo (SINTÉTICO/MINERAL)

## How to run

```powershell
# Activate virtual environment
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install streamlit azure-ai-formrecognizer openpyxl azure-core pandas pymupdf pillow numpy

# Run
streamlit run app.py
```

Credentials are loaded from `.streamlit/secrets.toml` (Streamlit secrets mechanism).

## Configuration

- `.streamlit/secrets.toml` — contains three keys:
  - `AZURE_DI_ENDPOINT` — Azure Document Intelligence endpoint URL
  - `AZURE_DI_KEY` — API key
  - `AZURE_DI_MODEL_ID` — custom model name (`oil-card3`)
- Do **not** commit `secrets.toml` to version control

## Change history

### 2026-03-09
- Initial context file created; no code changes made this session

## Known issues / things to verify

- The custom model (`oil-card3`) returns `None` for all `selectionMark` fields on scanned forms — the page-marks fallback was built specifically to work around this Azure limitation
- Page pre-filter thresholds (`left_mean < 185`, `left_std > 25`) are tuned for this specific form template; may need re-tuning for significantly different scan quality or layouts
- Fuzzy cutoff for compartment word matching drops to 0.45 for individual words (to handle barcode sticker OCR corruption) — could generate false positives in edge cases
- The `requirements.txt` is not a standard pip requirements file — it contains PowerShell commands and env var setup notes; no `pip freeze` output exists

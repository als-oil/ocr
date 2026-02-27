import re
import unicodedata
from io import BytesIO
from difflib import get_close_matches

import streamlit as st

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential


# ============================================================
# 0) DEPENDENCIES FOR PDF PAGE FILTERING
# ============================================================
# requirements.txt (add these):
# pymupdf
# pillow
# numpy


# ============================================================
# 1) EXCEL TEMPLATE (15 COLUMNS) + STYLE
# ============================================================
COLETA_COLUMNS = [
    "Chassi Série",
    "Tag Frota",
    "Ponto de Coleta / Compartimento",
    "Horímetro/Km/Período",
    "Número do Frasco",
    "Data da Coleta",
    "Óleo trocado",
    "Volume adicionado",
    "Fabricante (Óleo)",   # must exist but MUST be empty (manual fill)
    "Viscosidade (Óleo)",
    "Modelo (Óleo)",       # must exist but MUST be empty (manual fill)
    "Descrição do Óleo",
    "Horas/Km do Fluído",
    "Comentário",
    "Código externo",
]

FORCE_EMPTY_IN_EXCEL = {"Fabricante (Óleo)", "Modelo (Óleo)"}

COL_WIDTHS = {
    "Chassi Série": 22,
    "Tag Frota": 22,
    "Ponto de Coleta / Compartimento": 32,
    "Horímetro/Km/Período": 25,
    "Número do Frasco": 25,
    "Data da Coleta": 16,
    "Óleo trocado": 18,
    "Volume adicionado": 22,
    "Fabricante (Óleo)": 20,
    "Viscosidade (Óleo)": 20,
    "Modelo (Óleo)": 18,
    "Descrição do Óleo": 22,
    "Horas/Km do Fluído": 22,
    "Comentário": 22,
    "Código externo": 18,
}

GREEN_FILL = PatternFill("solid", fgColor="D9EAD3")
PINK_FILL = PatternFill("solid", fgColor="F4CCCC")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")

GREEN_COLS = {
    "Horímetro/Km/Período",
    "Data da Coleta",
    "Óleo trocado",
    "Volume adicionado",
    "Modelo (Óleo)",
    "Descrição do Óleo",
    "Horas/Km do Fluído",
    "Código externo",
}

PINK_COLS = {
    "Chassi Série",
    "Tag Frota",
    "Ponto de Coleta / Compartimento",
    "Número do Frasco",
    "Fabricante (Óleo)",
    "Viscosidade (Óleo)",
    "Comentário",
}

HEADER_FONT = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

thin = Side(border_style="thin", color="999999")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# Drop-down for "Óleo trocado"
OIL_TROCADO_OPTIONS = ["Sim", "Não", ""]


def create_template_workbook() -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Coleta"

    # Header row
    ws.append(COLETA_COLUMNS)
    ws.row_dimensions[1].height = 28

    for col_idx, col_name in enumerate(COLETA_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER

        # Fill
        if col_name in GREEN_COLS:
            cell.fill = GREEN_FILL
        elif col_name in PINK_COLS:
            cell.fill = PINK_FILL
        else:
            cell.fill = WHITE_FILL

        # Column widths
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(col_name, 18)

    # Data Validation for Óleo trocado
    oil_col = COLETA_COLUMNS.index("Óleo trocado") + 1
    dv = DataValidation(type="list", formula1=f'"{",".join(OIL_TROCADO_OPTIONS)}"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{get_column_letter(oil_col)}2:{get_column_letter(oil_col)}5000")

    return wb


def write_records_to_workbook(wb: openpyxl.Workbook, records: list[dict]) -> openpyxl.Workbook:
    ws = wb["Coleta"]

    for rec in records:
        row = [rec.get(c, "") for c in COLETA_COLUMNS]
        ws.append(row)

    # Styling for data rows
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 24
        for c_idx, col_name in enumerate(COLETA_COLUMNS, start=1):
            cell = ws.cell(row=r, column=c_idx)
            cell.alignment = LEFT
            cell.border = BORDER

            # Fill pattern like header
            if col_name in GREEN_COLS:
                cell.fill = GREEN_FILL
            elif col_name in PINK_COLS:
                cell.fill = PINK_FILL
            else:
                cell.fill = WHITE_FILL

    return wb


# ============================================================
# 2) NORMALIZATION + FIELD MAPPING
# ============================================================
def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 /?]", "", s)
    return s.strip()


COL_NORM_MAP = {_norm(c): c for c in COLETA_COLUMNS}
COL_NORMS = list(COL_NORM_MAP.keys())

SYNONYMS = {
    # Chassi
    "seriechassi": "Chassi Série",
    "serie/chassi": "Chassi Série",
    "serie chassi": "Chassi Série",
    "serie do chassi": "Chassi Série",
    "chassi serie": "Chassi Série",
    "chasis": "Chassi Série",
    "serie/chasis": "Chassi Série",

    # Tag frota
    "frota/tag": "Tag Frota",
    "frota tag": "Tag Frota",
    "tag frota": "Tag Frota",
    "tag": "Tag Frota",
    "equipo": "Tag Frota",
    "equipo tag": "Tag Frota",
    "flota": "Tag Frota",
    "flota/tag": "Tag Frota",

    # Ponto coleta
    "ponto de coleta": "Ponto de Coleta / Compartimento",
    "ponto de coleta compartimento": "Ponto de Coleta / Compartimento",
    "ponto coleta": "Ponto de Coleta / Compartimento",
    "compartimento": "Ponto de Coleta / Compartimento",
    "tipo de compartimento": "Ponto de Coleta / Compartimento",
    "tipo compartimento": "Ponto de Coleta / Compartimento",

    # Horimetro
    "horimetro/km/periodo": "Horímetro/Km/Período",
    "horimetro km periodo": "Horímetro/Km/Período",
    "horimetro": "Horímetro/Km/Período",
    "horometro": "Horímetro/Km/Período",
    "km/periodo": "Horímetro/Km/Período",

    # Óleo / Fluido trocado
    "oleo trocado": "Óleo trocado",
    "óleo trocado": "Óleo trocado",
    "fluido trocado": "Óleo trocado",
    "fluído trocado": "Óleo trocado",
    "fluido trocado?": "Óleo trocado",
    "fluído trocado?": "Óleo trocado",
    "aceite cambiado": "Óleo trocado",
    "aceite cambiado?": "Óleo trocado",

    # Número do frasco
    "amostra": "Número do Frasco",
    "muestra": "Número do Frasco",

    # Código externo
    "codigo ext/os": "Código externo",
    "codigo ext./os": "Código externo",
    "codigo ext os": "Código externo",
    "codigo ext/ot": "Código externo",
    "codigo ext./ot": "Código externo",
    "codigo ext ot": "Código externo",
    "codigo externo": "Código externo",

    # Data
    "data da coleta": "Data da Coleta",
    "data coleta": "Data da Coleta",
    "fecha de muestreo": "Data da Coleta",
    "fecha muestreo": "Data da Coleta",
    "fecha de muestra": "Data da Coleta",
    "fecha": "Data da Coleta",

    # Volume adicionado
    "vol oleo adic": "Volume adicionado",
    "vol. oleo adic": "Volume adicionado",
    "vol fluido adic": "Volume adicionado",
    "vol. fluido adic": "Volume adicionado",
    "vol fluido adicionado": "Volume adicionado",
    "vol. fluido adicionado": "Volume adicionado",

    # Viscosidade
    "viscosidade": "Viscosidade (Óleo)",
    "viscosidade oleo": "Viscosidade (Óleo)",
    "viscosidad": "Viscosidade (Óleo)",

    # Descrição do óleo
    "fabricante e modelo": "Descrição do Óleo",
    "fabricante y modelo": "Descrição do Óleo",
    "descripcion del aceite": "Descrição do Óleo",
    "descricao do oleo": "Descrição do Óleo",
    "descricao do óleo": "Descrição do Óleo",
    "descricao": "Descrição do Óleo",

    # Horas/km do fluído
    "horas/km do fluido": "Horas/Km do Fluído",
    "horas km do fluido": "Horas/Km do Fluído",
    "horas/km de aceite": "Horas/Km do Fluído",
    "horas km de aceite": "Horas/Km do Fluído",
    "horas/km do oleo": "Horas/Km do Fluído",
    "horas km do oleo": "Horas/Km do Fluído",

    # Comentário
    "observacoes/feedback": "Comentário",
    "observacoes": "Comentário",
    "observaciones": "Comentário",
    "comentario": "Comentário",
    "comentarios": "Comentário",

    # Even if your model emits these, we force blank in output:
    "fabricante oleo": "Fabricante (Óleo)",
    "fabricante (oleo)": "Fabricante (Óleo)",
    "modelo": "Modelo (Óleo)",
    "modelo oleo": "Modelo (Óleo)",
    "modelo (oleo)": "Modelo (Óleo)",
}


def _empty_if_none_like(v: str) -> str:
    vv = (v or "").strip()
    if vv.lower() in ("none", "null", "nan", "-", "n/a", "unselected", "unreadable", "illegible"):
        return ""
    return vv


def _expand_two_digit_year(yy: str) -> str:
    yy = yy.strip()
    if len(yy) == 2 and yy.isdigit():
        return "20" + yy
    return yy


def _normalize_date_str(v: str) -> str:
    """
    dd/mm/yy -> dd/mm/yyyy
    dd-mm-yy -> dd/mm/yyyy
    and if value is ONLY '24' -> '2024'
    """
    v = (v or "").strip()
    if not v:
        return ""

    if re.fullmatch(r"\d{2}", v):
        return _expand_two_digit_year(v)

    v2 = v.replace("-", "/")
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", v2)
    if not m:
        return v

    dd, mm, yy = m.group(1).zfill(2), m.group(2).zfill(2), m.group(3)
    yy = _expand_two_digit_year(yy)
    return f"{dd}/{mm}/{yy}"


def clean_value(col_name: str, v: str) -> str:
    v = _empty_if_none_like(v)
    if not v:
        return ""

    if col_name in FORCE_EMPTY_IN_EXCEL:
        return ""

    if col_name == "Data da Coleta":
        return _normalize_date_str(v)

    if col_name in ("Número do Frasco", "Código externo"):
        # prefer long numeric token (barcode)
        m = re.search(r"\b(\d{6,})\b", v)
        return m.group(1) if m else v

    if col_name == "Óleo trocado":
        vn = _norm(v)
        if vn in ("sim", "si", "yes", "selected", "true", "1", "x"):
            return "Sim"
        if vn in ("nao", "não", "no"):
            return "Não"

    return v


def key_to_column(k: str) -> str | None:
    kn = _norm(k)

    # Heuristic: some models emit "Horas/Km do Fluído" with unexpected variations.
    if (("horas" in kn) or ("hora" in kn) or ("km" in kn)) and (("fluido" in kn) or ("oleo" in kn) or ("aceite" in kn)):
        return "Horas/Km do Fluído"

    if kn in SYNONYMS:
        return SYNONYMS[kn]
    if kn in COL_NORM_MAP:
        return COL_NORM_MAP[kn]

    best = get_close_matches(kn, COL_NORMS, n=1, cutoff=0.78)
    if best:
        return COL_NORM_MAP[best[0]]
    return None


def fields_to_record(fields: dict) -> dict:
    record = {c: "" for c in COLETA_COLUMNS}

    def is_selected(v) -> bool:
        return str(v).strip().lower() in ("selected", "sim", "si", "yes", "true", "1", "x")

    for k, v in (fields or {}).items():
        if v is None:
            continue
        v_str = str(v).strip()
        if _empty_if_none_like(v_str) == "":
            continue

        kn = _norm(k)

        # checkbox-like "óleo/fluido trocado"
        # Models may emit:
        # - a single field with value "Sim/Não"
        # - two option fields (Sim/Não) with value "selected/true/x"
        # - keys without separators (e.g., "...Sim")
        if ("oleo trocado" in kn) or ("fluido trocado" in kn) or ("aceite cambiado" in kn) or ("fluido trocado" in kn.replace(" ", "")):

            v_clean = _norm(v_str)

            def _yesno_from_token(s: str) -> str | None:
                # kn/v_clean are already normalized (no accents)
                if re.search(r"(sim|si|yes)$", s) or re.search(r"\b(sim|si|yes)\b", s):
                    return "Sim"
                if re.search(r"(nao|no)$", s) or re.search(r"\b(nao|no)\b", s):
                    return "Não"
                return None

            # Case A: value itself already indicates yes/no
            yn = _yesno_from_token(v_clean)
            if yn:
                record["Óleo trocado"] = yn
                continue

            # Case B: key indicates the option and value indicates selection
            if is_selected(v_str):
                ynk = _yesno_from_token(kn.replace(" ", ""))
                if ynk:
                    record["Óleo trocado"] = ynk
                    continue

            # Case C: free-text fallback
            if "sim" in v_clean or "si" in v_clean or "yes" in v_clean:
                record["Óleo trocado"] = "Sim"
                continue
            if "nao" in v_clean or "no" in v_clean:
                record["Óleo trocado"] = "Não"
                continue

            continue

        col = key_to_column(k)
        if not col:
            continue

        record[col] = clean_value(col, v_str)

    # Always blank these (manual fill)
    for c in FORCE_EMPTY_IN_EXCEL:
        record[c] = ""

    return record


# ============================================================
# 3) FILTERING: IGNORE NON-FORM PAGES (PDFS WITH MIXED CONTENT)
# ============================================================
def _record_has_signal(rec: dict) -> bool:
    """
    Safety net: keep only records that look like real forms.
    Strong signals:
    - Número do Frasco or Código externo has >= 6 digits
    OR
    - at least 3 key fields filled
    """
    def has_big_number(s: str) -> bool:
        return bool(re.search(r"\b\d{6,}\b", (s or "")))

    if has_big_number(rec.get("Número do Frasco", "")):
        return True
    if has_big_number(rec.get("Código externo", "")):
        return True

    key_fields = [
        "Chassi Série",
        "Tag Frota",
        "Ponto de Coleta / Compartimento",
        "Horímetro/Km/Período",
        "Data da Coleta",
        "Óleo trocado",
        "Viscosidade (Óleo)",
        "Descrição do Óleo",
        "Horas/Km do Fluído",
    ]
    filled = sum(1 for k in key_fields if (rec.get(k, "") or "").strip())
    return filled >= 3


def _looks_like_form_by_fields(raw_fields: dict) -> bool:
    """
    Decide if a page is a form based on which fields the model detected.

    The idea:
    - If the model extracted "enough" of the expected form field keys, it's a form page.
    """
    if not raw_fields:
        return False

    expected_signals = [
        "serie", "chassi", "tag", "frota", "compartimento", "ponto", "coleta",
        "horimetro", "horometro", "periodo", "amostra", "muestra", "codigo",
        "data", "fecha", "viscos", "descricao", "aceite", "oleo", "fluido",
        "horas", "km"
    ]

    keys_norm = [_norm(k) for k in raw_fields.keys()]
    matches = 0
    for kn in keys_norm:
        if any(sig in kn for sig in expected_signals):
            matches += 1

    # Threshold: tuneable. 3 is a good starting point for "at least part of the form"
    return matches >= 3


# ============================================================
# 4) AZURE DOCUMENT INTELLIGENCE
# ============================================================
ENDPOINT = st.secrets["AZURE_DI_ENDPOINT"]
KEY = st.secrets["AZURE_DI_KEY"]
MODEL_ID = st.secrets["AZURE_DI_MODEL_ID"]

di_client = DocumentAnalysisClient(ENDPOINT, AzureKeyCredential(KEY))


def analyze_bytes(file_bytes: bytes):
    poller = di_client.begin_analyze_document(MODEL_ID, document=file_bytes)
    return poller.result()


def result_to_records(result) -> list[dict]:
    records: list[dict] = []
    if getattr(result, "documents", None):
        for doc in result.documents:
            fields = {}
            for name, field in (doc.fields or {}).items():
                fields[name] = field.value if field.value is not None else field.content
            records.append(fields_to_record(fields))
    return records


# ============================================================
# 4b) PDF SPLITTING (FOR MIXED PDFs) — ONLY KEEP FORM PAGES
# ============================================================
def split_pdf_into_pages(pdf_bytes: bytes) -> list[bytes]:
    import fitz  # pymupdf

    pages = []
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    for i in range(src.page_count):
        doc = fitz.open()
        doc.insert_pdf(src, from_page=i, to_page=i)
        pages.append(doc.tobytes())
        doc.close()
    src.close()
    return pages


def extract_records_from_upload(file_bytes: bytes, mime: str) -> tuple[list[dict], dict]:
    """
    Returns (records, stats)
    stats example: {"total_pages": 10, "kept_pages": 6, "dropped_pages": 4}
    """
    stats = {"total_pages": 1, "kept_pages": 0, "dropped_pages": 0}

    # PDF: split into pages and analyze each page (ignore non-form pages)
    if mime == "application/pdf":
        pages = split_pdf_into_pages(file_bytes)
        stats["total_pages"] = len(pages)

        records: list[dict] = []
        for pbytes in pages:
            result = analyze_bytes(pbytes)

            # Keep page if it "looks like a form" by detected fields OR by filled record signals
            raw_fields = {}
            if getattr(result, "documents", None):
                for doc in result.documents:
                    for name, field in (doc.fields or {}).items():
                        raw_fields[name] = field.value if field.value is not None else field.content

            page_records = result_to_records(result)
            keep = _looks_like_form_by_fields(raw_fields) or any(_record_has_signal(r) for r in page_records)

            if keep:
                records.extend([r for r in page_records if _record_has_signal(r)])
                stats["kept_pages"] += 1
            else:
                stats["dropped_pages"] += 1

        return records, stats

    # Image (jpg/png/etc): just analyze once
    result = analyze_bytes(file_bytes)
    records = [r for r in result_to_records(result) if _record_has_signal(r)]
    stats["kept_pages"] = 1 if records else 0
    stats["dropped_pages"] = 0 if records else 1
    return records, stats


# ============================================================
# 5) STREAMLIT UI (ONLY: UPLOAD -> EXTRACT -> DOWNLOAD EXCEL)
# ============================================================
st.title("OCR – Cartão de Óleo → Excel")
st.caption("Fluxo: enviar arquivo → extrair → baixar Excel. (Ignora páginas que não são formulário.)")

uploaded_files = st.file_uploader(
    "Envie 1+ cartões (imagens) ou PDFs",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_records = []
    total_stats = {"total_pages": 0, "kept_pages": 0, "dropped_pages": 0}

    for uf in uploaded_files:
        file_bytes = uf.getvalue()
        mime = uf.type

        records, stats = extract_records_from_upload(file_bytes, mime)

        all_records.extend(records)
        total_stats["total_pages"] += stats.get("total_pages", 1)
        total_stats["kept_pages"] += stats.get("kept_pages", 0)
        total_stats["dropped_pages"] += stats.get("dropped_pages", 0)

    if not all_records:
        st.warning("Nenhum formulário identificado nos arquivos enviados.")
    else:
        wb = create_template_workbook()
        wb = write_records_to_workbook(wb, all_records)

        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success(
            f"Extração concluída. Registros: {len(all_records)} | "
            f"Páginas/arquivos: {total_stats['total_pages']} | "
            f"Mantidas: {total_stats['kept_pages']} | "
            f"Ignoradas: {total_stats['dropped_pages']}"
        )

        st.download_button(
            "Baixar Excel",
            data=out,
            file_name="coleta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

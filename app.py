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
    "Comentário",
    "Código externo",
}
PINK_COLS = {
    "Número do Frasco",
    "Fabricante (Óleo)",
    "Viscosidade (Óleo)",
}

DV_PONTO_COLETA = ["MOTOR", "REDUTOR", "TRANSMISSÃO", "DIFERENCIAL", "HIDRÁULICO", "COMPRESSOR", "RADIADOR", "OUTROS"]
DV_OLEO_TROCADO = ["Sim", "Não"]
DV_DESCRICAO = ["SINTÉTICO", "MINERAL"]

CENTER_COLS = {
    "Número do Frasco",
    "Código externo",
    "Data da Coleta",
    "Óleo trocado",
    "Horímetro/Km/Período",
}


def build_excel_bytes(records: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Coleta"
    wb.create_sheet("Referências")

    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center")

    # Header
    for col_idx, col_name in enumerate(COLETA_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.alignment = header_align
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(col_name, 18)
    ws.row_dimensions[1].height = 22

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLETA_COLUMNS))}1"

    # Data rows (force empty manual columns)
    for r_idx, rec in enumerate(records, start=2):
        for c_idx, col_name in enumerate(COLETA_COLUMNS, start=1):
            value = "" if col_name in FORCE_EMPTY_IN_EXCEL else rec.get(col_name, "")
            ws.cell(row=r_idx, column=c_idx, value=value)

    max_row = max(2, ws.max_row)
    max_col = len(COLETA_COLUMNS)

    # Fills + alignment
    for c_idx, col_name in enumerate(COLETA_COLUMNS, start=1):
        fill = GREEN_FILL if col_name in GREEN_COLS else PINK_FILL if col_name in PINK_COLS else WHITE_FILL
        for r_idx in range(1, max_row + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.fill = fill
            if r_idx >= 2 and col_name in CENTER_COLS:
                cell.alignment = center_align

    # Borders (grid + strong vertical separators)
    thin = Side(style="thin", color="B7B7B7")
    thick = Side(style="medium", color="808080")

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        for r in range(1, max_row + 1):
            cell = ws[f"{col_letter}{r}"]
            cell.border = Border(
                left=cell.border.left,
                right=thick,
                top=cell.border.top,
                bottom=cell.border.bottom,
            )

    # Dropdown validations (apply down to 500 rows)
    last_row = max(500, max_row)

    def add_list_validation(col_name: str, options: list[str]):
        if col_name not in COLETA_COLUMNS:
            return
        col_idx = COLETA_COLUMNS.index(col_name) + 1
        col_letter = get_column_letter(col_idx)
        dv = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=True, showDropDown=True)
        ws.add_data_validation(dv)
        dv.add(f"{col_letter}2:{col_letter}{last_row}")

    add_list_validation("Ponto de Coleta / Compartimento", DV_PONTO_COLETA)
    add_list_validation("Óleo trocado", DV_OLEO_TROCADO)
    add_list_validation("Descrição do Óleo", DV_DESCRICAO)

    buf = BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


# ============================================================
# 2) NORMALIZATION + MAPPING (OCR KEYS -> EXCEL COLUMNS)
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
    v = _empty_if_none_like(v)
    if not v:
        return ""

    vv = v.strip()

    # Only "24" -> "2024"
    if re.fullmatch(r"\d{2}", vv):
        return _expand_two_digit_year(vv)

    # dd/mm/yy or dd-mm-yy or dd/mm/yyyy
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", vv)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        y = _expand_two_digit_year(y)
        return f"{int(d):02d}/{int(mo):02d}/{y}"

    # yyyy-mm-dd -> dd/mm/yyyy
    m2 = re.search(r"\b(\d{4})[/-](\d{1,2})[/-](\d{1,2})\b", vv)
    if m2:
        y, mo, d = m2.group(1), m2.group(2), m2.group(3)
        return f"{int(d):02d}/{int(mo):02d}/{y}"

    return vv


def clean_value(col_name: str, value: str) -> str:
    v = _empty_if_none_like(value)
    if not v:
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

        # checkbox-like "óleo/fluido trocado sim/não"
                # checkbox-like "óleo/fluido trocado"
        # Handles:
        # - "Fluido trocado: Sim"
        # - "Fluido trocado Sim" = selected
        # - "Óleo trocado Não" = selected
        # - "Aceite cambiado: SI"
        if ("oleo trocado" in kn) or ("fluido trocado" in kn) or ("aceite cambiado" in kn):
            v_clean = _norm(v_str)

            # Case A: value already sim/nao
            if v_clean in ("sim", "si", "yes", "true", "1", "x"):
                record["Óleo trocado"] = "Sim"
                continue
            if v_clean in ("nao", "não", "no"):
                record["Óleo trocado"] = "Não"
                continue

            # Case B: key contains Sim/Não and value is "selected"
            if is_selected(v_str):
                if (" sim" in kn) or kn.endswith(" sim") or (" si" in kn) or kn.endswith(" si"):
                    record["Óleo trocado"] = "Sim"
                    continue
                if (" nao" in kn) or (" no" in kn) or kn.endswith(" nao") or kn.endswith(" no"):
                    record["Óleo trocado"] = "Não"
                    continue

            # Case C: fallback – sim/nao inside free text
            if "sim" in v_clean:
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


def _is_probably_form_page(pil_img) -> bool:
    """
    Cheap pre-filter (no Azure call):
    Form pages have a large darker/grey block on the left half.
    Non-form pages (cover, preregistration, blank, etc) are more uniformly light.
    """
    import numpy as np
    from PIL import Image

    img = pil_img.convert("L")
    img = img.resize((600, int(600 * img.height / img.width)), Image.BILINEAR)

    arr = np.array(img, dtype=np.float32)
    h, w = arr.shape

    left = arr[:, : w // 2]
    right = arr[:, w // 2 :]

    left_mean = float(left.mean())
    left_std = float(left.std())
    right_mean = float(right.mean())

    # tuned for this template type
    looks_like_form = (left_mean < 185) and (left_std > 25)
    not_blank = (left_mean < 245) or (right_mean < 245)

    return bool(looks_like_form and not_blank)


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


def pdf_split_to_single_page_pdfs(pdf_bytes: bytes) -> list[bytes]:
    """
    Splits multi-page PDF into single-page PDF bytes (no rasterization, avoids huge images).
    Requires: pymupdf
    """
    import fitz  # PyMuPDF

    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out: list[bytes] = []
    for i in range(src.page_count):
        dst = fitz.open()
        dst.insert_pdf(src, from_page=i, to_page=i)
        out.append(dst.tobytes())
        dst.close()
    src.close()
    return out


def pdf_single_page_pdf_to_pil(single_page_pdf_bytes: bytes, dpi: int = 72):
    """
    Renders a single-page PDF to a small PIL image only for classification.
    Low DPI keeps it fast and avoids size limits.
    Requires: pymupdf, pillow
    """
    import fitz
    from PIL import Image
    import io

    doc = fitz.open(stream=single_page_pdf_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    doc.close()

    return Image.open(io.BytesIO(pix.tobytes("png")))


def extract_records_from_upload(file_bytes: bytes, mime_type: str) -> tuple[list[dict], dict]:
    """
    Returns (records, stats).
    - Images: analyze once.
    - PDFs: split into single-page PDFs; pre-filter pages (form-like) to avoid OCR on non-forms;
            then analyze only candidates; post-filter records as a safety net.
    """
    stats = {
        "pdf_pages_total": 0,
        "pdf_pages_candidates": 0,
        "records_before_filter": 0,
        "records_after_filter": 0,
        "skipped_pages": 0,
    }

    # Non-PDF: single call
    if mime_type != "application/pdf":
        result = analyze_bytes(file_bytes)
        records = result_to_records(result)
        stats["records_before_filter"] = len(records)
        records = [r for r in records if _record_has_signal(r)]
        stats["records_after_filter"] = len(records)
        return records, stats

    # PDF: split pages
    page_pdfs = pdf_split_to_single_page_pdfs(file_bytes)
    stats["pdf_pages_total"] = len(page_pdfs)

    # Pre-filter (local classifier)
    candidate_pages: list[bytes] = []
    for spdf in page_pdfs:
        try:
            pil = pdf_single_page_pdf_to_pil(spdf, dpi=72)
            if _is_probably_form_page(pil):
                candidate_pages.append(spdf)
            else:
                stats["skipped_pages"] += 1
        except Exception:
            # failsafe: if classification fails, keep page (better than losing data)
            candidate_pages.append(spdf)

    stats["pdf_pages_candidates"] = len(candidate_pages)

    # OCR only candidate pages
    all_records: list[dict] = []
    for spdf in candidate_pages:
        result = analyze_bytes(spdf)
        all_records.extend(result_to_records(result))

    stats["records_before_filter"] = len(all_records)

    # Post-filter (safety net)
    cleaned = [r for r in all_records if _record_has_signal(r)]
    stats["records_after_filter"] = len(cleaned)

    return cleaned, stats


# ============================================================
# 5) STREAMLIT UI (ONLY: UPLOAD -> EXTRACT -> DOWNLOAD EXCEL)
# ============================================================
st.title("OCR – Cartão de Óleo → Excel")
st.caption("Fluxo: enviar arquivo → extrair → baixar Excel. (Ignora páginas que não são formulário.)")

uploaded_file = st.file_uploader(
    "Envie um cartão (imagem) ou um PDF com várias páginas (formulários misturados)",
    type=["jpg", "jpeg", "png", "pdf"],
)

if uploaded_file is not None and uploaded_file.type.startswith("image/"):
    st.image(uploaded_file, caption="Imagem enviada")

if uploaded_file is None:
    st.info("Envie um arquivo para habilitar a extração.")
    st.stop()

st.markdown("### Extrair e gerar Excel")

if st.button("🚀 Extrair somente formulários e baixar Excel"):
    with st.spinner("Processando..."):
        try:
            records, stats = extract_records_from_upload(uploaded_file.getvalue(), uploaded_file.type)
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
            st.stop()

    if uploaded_file.type == "application/pdf":
        st.write(
            f"PDF: {stats['pdf_pages_total']} páginas | "
            f"candidatas (form): {stats['pdf_pages_candidates']} | "
            f"ignoradas: {stats['skipped_pages']}"
        )

    st.write(f"Registros (antes do filtro): {stats['records_before_filter']}")
    st.write(f"Registros (após filtro): **{stats['records_after_filter']}**")

    if not records:
        st.warning("Não consegui extrair registros úteis (ou todas as páginas foram classificadas como não-formulário).")
        st.stop()

    with st.expander("Prévia (primeiras 5 linhas)"):
        for i, rec in enumerate(records[:5], start=1):
            st.markdown(f"**Linha {i}**")
            for c in COLETA_COLUMNS:
                st.write(f"- {c}: {rec.get(c, '')}")
            st.divider()

    excel_bytes = build_excel_bytes(records)
    st.download_button(
        "📥 Baixar Excel",
        data=excel_bytes,
        file_name="coleta.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


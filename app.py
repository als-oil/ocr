import json
import datetime
from pathlib import Path
from io import BytesIO
import re
import unicodedata
from difflib import get_close_matches

import streamlit as st

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential


# ============================================================
# 1) EXCEL OUTPUT STRUCTURE (DEFAULT FORMAT)
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
    "Fabricante (Óleo)",
    "Viscosidade (Óleo)",
    "Modelo (Óleo)",
    "Descrição do Óleo",
    "Horas/Km do Fluído",
    "Comentário",
    "Código externo",
]

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
PINK_FILL  = PatternFill("solid", fgColor="F4CCCC")
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
    """Build a styled Excel with filters, frozen header, widths, colors, borders and dropdown validations."""
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

    # Data rows
    for r_idx, rec in enumerate(records, start=2):
        for c_idx, col_name in enumerate(COLETA_COLUMNS, start=1):
            ws.cell(row=r_idx, column=c_idx, value=rec.get(col_name, ""))

    max_row = max(2, ws.max_row)
    max_col = len(COLETA_COLUMNS)

    # Column fills + center alignment where needed
    for c_idx, col_name in enumerate(COLETA_COLUMNS, start=1):
        fill = GREEN_FILL if col_name in GREEN_COLS else PINK_FILL if col_name in PINK_COLS else WHITE_FILL

        for r_idx in range(1, max_row + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.fill = fill
            if r_idx >= 2 and col_name in CENTER_COLS:
                cell.alignment = center_align

    # Borders (gridlines + strong vertical separators)
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

    # Dropdown validations
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
# 2) NOTES -> INTERNAL JSON (PARSE EDITABLE NOTES)
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
    "seriechassi": "Chassi Série",
    "serie/chassi": "Chassi Série",
    "serie/chassi*": "Chassi Série",
    "serie chassi": "Chassi Série",
    "serie do chassi": "Chassi Série",
    "chassi serie": "Chassi Série",
    "chasis": "Chassi Série",
    "serie/chasis": "Chassi Série",

    "frota/tag": "Tag Frota",
    "frota tag": "Tag Frota",
    "tag frota": "Tag Frota",
    "tag": "Tag Frota",
    "equipo": "Tag Frota",
    "equipo tag": "Tag Frota",
    "flota": "Tag Frota",
    "flota/tag": "Tag Frota",

    "ponto de coleta": "Ponto de Coleta / Compartimento",
    "ponto de coleta compartimento": "Ponto de Coleta / Compartimento",
    "ponto coleta": "Ponto de Coleta / Compartimento",
    "compartimento": "Ponto de Coleta / Compartimento",
    "tipo de compartimento": "Ponto de Coleta / Compartimento",
    "tipo compartimento": "Ponto de Coleta / Compartimento",

    "horimetro/km/periodo": "Horímetro/Km/Período",
    "horimetro km periodo": "Horímetro/Km/Período",
    "horimetro": "Horímetro/Km/Período",
    "km/periodo": "Horímetro/Km/Período",
    "horometrokmperiodo": "Horímetro/Km/Período",
    "horometro/km/periodo": "Horímetro/Km/Período",
    "horometro": "Horímetro/Km/Período",

    # IMPORTANT: Amostra -> Número do Frasco
    "amostra": "Número do Frasco",
    "muestra": "Número do Frasco",

    # IMPORTANT: Código Ext -> Código externo
    "codigo ext/os": "Código externo",
    "codigo ext./os": "Código externo",
    "codigo ext os": "Código externo",
    "codigo ext/ot": "Código externo",
    "codigo ext./ot": "Código externo",
    "codigo ext ot": "Código externo",
    "codigo externo": "Código externo",

    "data da coleta": "Data da Coleta",
    "data coleta": "Data da Coleta",
    "fecha de muestreo": "Data da Coleta",
    "fecha muestreo": "Data da Coleta",
    "fecha de muestra": "Data da Coleta",
    "fecha": "Data da Coleta",

    "vol oleo adic": "Volume adicionado",
    "vol. oleo adic": "Volume adicionado",
    "vol oleo adic.": "Volume adicionado",
    "vol. oleo adic.": "Volume adicionado",
    "vol fluido adic": "Volume adicionado",
    "vol. fluido adic": "Volume adicionado",
    "vol fluido adic.": "Volume adicionado",
    "vol. fluido adic.": "Volume adicionado",
    "vol fluido adicionado": "Volume adicionado",
    "vol. fluido adicionado": "Volume adicionado",

    "viscosidade": "Viscosidade (Óleo)",
    "viscosidade*": "Viscosidade (Óleo)",
    "viscosidade oleo": "Viscosidade (Óleo)",
    "viscosidad": "Viscosidade (Óleo)",

    "fabricante e modelo": "Descrição do Óleo",
    "fabricante e modelo*": "Descrição do Óleo",
    "fabricante y modelo de aceite": "Descrição do Óleo",
    "fabricante y modelo": "Descrição do Óleo",
    "descripcion del aceite": "Descrição do Óleo",
    "descricao do oleo": "Descrição do Óleo",
    "descricao do óleo": "Descrição do Óleo",
    "descricao oleo": "Descrição do Óleo",
    "descricao": "Descrição do Óleo",

    "modelo": "Modelo (Óleo)",
    "modelo oleo": "Modelo (Óleo)",
    "modelo (oleo)": "Modelo (Óleo)",

    "fabricante oleo": "Fabricante (Óleo)",
    "fabricante (oleo)": "Fabricante (Óleo)",

    "horas/km do fluido": "Horas/Km do Fluído",
    "horas km do fluido": "Horas/Km do Fluído",
    "horas/km de aceite": "Horas/Km do Fluído",
    "horas km de aceite": "Horas/Km do Fluído",
    "horas/km do oleo": "Horas/Km do Fluído",
    "horas km do oleo": "Horas/Km do Fluído",

    "observacoes/feedback": "Comentário",
    "observacoes": "Comentário",
    "observacoes feedback": "Comentário",
    "observaciones/feedback": "Comentário",
    "observacion/feedback": "Comentário",
    "observacion": "Comentário",
    "observaciones": "Comentário",
    "comentario": "Comentário",
    "comentarios": "Comentário",
}


def _empty_if_none_like(v: str) -> str:
    vv = (v or "").strip()
    if vv.lower() in ("none", "null", "nan", "-", "n/a", "unselected"):
        return ""
    return vv


def _normalize_date_str(v: str) -> str:
    v = _empty_if_none_like(v)
    if not v:
        return ""
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", v)
    if not m:
        return v
    d, mo, y = m.group(1), m.group(2), m.group(3)
    if len(y) == 2:
        y = "20" + y
    return f"{int(d):02d}/{int(mo):02d}/{y}"


def clean_value(col_name: str, value: str) -> str:
    v = _empty_if_none_like(value)
    if not v:
        return ""

    if col_name == "Data da Coleta":
        return _normalize_date_str(v)

    if col_name in ("Número do Frasco", "Código externo"):
        m = re.search(r"\b(\d{6,})\b", v)
        return m.group(1) if m else v

    if col_name == "Óleo trocado":
        vn = _norm(v)
        if vn in ("sim", "si", "yes", "selected", "true", "1", "x"):
            return "Sim"
        if vn in ("nao", "no"):
            return "Não"

    return v


def parse_notepad_to_record(notepad_text: str) -> dict:
    record = {c: "" for c in COLETA_COLUMNS}

    for raw_line in (notepad_text or "").splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue

        k, v = line.split(":", 1)
        k_norm = _norm(k.strip())
        v = v.strip()

        # óleo/fluido trocado (checkbox style)
        if ("oleo trocado" in k_norm or "fluido trocado" in k_norm or "aceite cambiado" in k_norm):
            v_clean = _empty_if_none_like(v).lower()
            if v_clean in ("selected", "sim", "si", "yes", "x", "1", "true"):
                if "sim" in k_norm or "si" in k_norm:
                    record["Óleo trocado"] = "Sim"
                elif "nao" in k_norm or "no" in k_norm or "não" in k_norm:
                    record["Óleo trocado"] = "Não"
            continue

        if k_norm in SYNONYMS:
            col = SYNONYMS[k_norm]
            record[col] = clean_value(col, v)
            continue

        if k_norm in COL_NORM_MAP:
            col = COL_NORM_MAP[k_norm]
            record[col] = clean_value(col, v)
            continue

        best = get_close_matches(k_norm, COL_NORMS, n=1, cutoff=0.78)
        if best:
            col = COL_NORM_MAP[best[0]]
            record[col] = clean_value(col, v)

    return record


# ============================================================
# 3) STORAGE (LOCAL)
# ============================================================
CLIENTS_DIR = Path("data/clients")
CLIENTS_DIR.mkdir(parents=True, exist_ok=True)

def list_clients() -> list[str]:
    return sorted([p.name for p in CLIENTS_DIR.iterdir() if p.is_dir()])

def ensure_client_folder(client: str) -> Path:
    folder = CLIENTS_DIR / client
    folder.mkdir(parents=True, exist_ok=True)
    return folder

def append_record(client: str, notes_text: str) -> None:
    folder = ensure_client_folder(client)
    record = parse_notepad_to_record(notes_text)
    payload = {
        "_saved_at": datetime.datetime.utcnow().isoformat() + "Z",
        "_notes": notes_text,
        **record,
    }
    with (folder / "records.jsonl").open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")

def read_records(client: str) -> list[dict]:
    path = CLIENTS_DIR / client / "records.jsonl"
    if not path.exists():
        return []
    out = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out

def records_for_excel(records: list[dict]) -> list[dict]:
    return [{c: r.get(c, "") for c in COLETA_COLUMNS} for r in records]


# ============================================================
# 4) AZURE CONFIG (SECRETS)
# ============================================================
ENDPOINT = st.secrets["AZURE_DI_ENDPOINT"]
KEY = st.secrets["AZURE_DI_KEY"]
MODEL_ID = st.secrets["AZURE_DI_MODEL_ID"]
di_client = DocumentAnalysisClient(ENDPOINT, AzureKeyCredential(KEY))


# ============================================================
# 5) OCR + NOTES
# ============================================================
def extract_fields(file_bytes: bytes) -> dict:
    poller = di_client.begin_analyze_document(MODEL_ID, document=file_bytes)
    result = poller.result()
    if not result.documents:
        return {}
    doc = result.documents[0]
    out = {}
    for name, field in doc.fields.items():
        out[name] = field.value if field.value is not None else field.content
    return out


def build_notepad_text(fields: dict) -> str:
    """All available info, excluding None/unselected/unreadable; show only selected checkboxes."""
    def is_unavailable(v) -> bool:
        if v is None:
            return True
        vv = str(v).strip()
        if not vv:
            return True
        return vv.lower() in ("none", "null", "nan", "-", "n/a", "unselected", "unreadable", "illegible")

    def is_selected(v) -> bool:
        return str(v).strip().lower() in ("selected", "sim", "si", "yes", "true", "1", "x")

    lines = []
    for k, v in (fields or {}).items():
        if is_unavailable(v):
            continue

        k_norm = _norm(k)
        v_str = str(v).strip()
        v_low = v_str.lower()

        # generic checkbox
        if v_low == "selected":
            lines.append(f"{k}: selected")
            continue

        # óleo/fluido trocado (only if selected)
        if ("oleo trocado" in k_norm) or ("fluido trocado" in k_norm):
            if is_selected(v_str):
                if "sim" in k_norm or "si" in k_norm:
                    lines.append(f"{k}: Sim")
                elif "nao" in k_norm or "não" in k_norm or "no" in k_norm:
                    lines.append(f"{k}: Não")
                else:
                    lines.append(f"{k}: {v_str}")
            continue

        lines.append(f"{k}: {v_str}")

    lines.sort(key=lambda s: _norm(s.split(":", 1)[0]))
    return "\n".join(lines)


# ============================================================
# 6) UI
# ============================================================
st.title("OCR – Cartão de Óleo")

uploaded_file = st.file_uploader("Envie o arquivo do cartão", type=["jpg", "jpeg", "png", "pdf"])

st.session_state.setdefault("last_fields", None)
st.session_state.setdefault("notes_text", "")
st.session_state.setdefault("last_filename", "")

st.markdown("### 1) Extrair com OCR")

if uploaded_file is not None and uploaded_file.type.startswith("image/"):
    st.image(uploaded_file, caption="Imagem enviada")

if uploaded_file is not None and st.button("Extrair dados com OCR"):
    with st.spinner("Processando..."):
        try:
            fields = extract_fields(uploaded_file.getvalue())
        except Exception as e:
            st.error(f"Erro ao chamar o modelo: {e}")
            st.stop()

    if not fields:
        st.warning("Nenhum campo retornado.")
    else:
        st.session_state.last_fields = fields
        st.session_state.notes_text = build_notepad_text(fields)
        st.session_state.last_filename = uploaded_file.name

if st.session_state.last_fields:
    st.subheader("Bloco de notas (edite se precisar)")
    st.session_state.notes_text = st.text_area(
        "Notas (fonte para gerar Excel)",
        st.session_state.notes_text,
        height=280,
    )

    with st.expander("Prévia do que vai para o Excel (15 colunas)"):
        preview = parse_notepad_to_record(st.session_state.notes_text)
        for c in COLETA_COLUMNS:
            st.write(f"**{c}**: {preview.get(c, '')}")


st.markdown("### 2) Salvar extração em um cliente / Exportar 1 linha (não salva)")

if not st.session_state.last_fields:
    st.info("Primeiro faça a extração com OCR para poder salvar/exportar.")
else:
    clients = list_clients()
    choice = st.selectbox("Cliente", ["(novo...)"] + clients)

    client_name = st.text_input("Nome do novo cliente").strip() if choice == "(novo...)" else choice

    colA, colB = st.columns(2)

    with colA:
        if st.button("💾 Salvar esta extração"):
            if not client_name:
                st.warning("Informe o nome do cliente.")
            else:
                append_record(client_name, st.session_state.notes_text)
                st.success(f"Salvo em data/clients/{client_name}/records.jsonl")

    with colB:
        if st.button("📄 Gerar Excel (1 linha, não salva)"):
            record = parse_notepad_to_record(st.session_state.notes_text)
            excel_bytes = build_excel_bytes([record])
            st.download_button(
                "📥 Baixar Excel (1 linha)",
                data=excel_bytes,
                file_name="coleta_1_linha.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


st.markdown("### 3) Exportar Excel por cliente")

clients = list_clients()
if not clients:
    st.info("Nenhum cliente salvo ainda.")
else:
    client_to_export = st.selectbox("Cliente para exportar", clients, key="export_client")
    records = read_records(client_to_export)
    st.write(f"Registros: **{len(records)}**")

    if records and st.button("Gerar Excel (cliente)"):
        excel_bytes = build_excel_bytes(records_for_excel(records))
        st.download_button(
            "📥 Baixar Excel do cliente",
            data=excel_bytes,
            file_name=f"{client_to_export}_coleta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

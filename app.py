import os

import streamlit as st
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# =========================
#  CONFIG AZURE
# =========================
ENDPOINT = os.getenv("AZURE_DI_ENDPOINT")
KEY = os.getenv("AZURE_DI_KEY")
MODEL_ID = "oil-card"  # ajuste se seu model_id for outro

if not ENDPOINT or not KEY:
    st.error(
        "AZURE_DI_ENDPOINT e/ou AZURE_DI_KEY não estão definidos.\n"
        "Defina as variáveis de ambiente antes de rodar o app."
    )
    st.stop()

client = DocumentAnalysisClient(ENDPOINT, AzureKeyCredential(KEY))


# =========================
#  FUNÇÕES
# =========================
def extract_fields(file_bytes: bytes) -> dict:
    """Chama o modelo customizado e retorna os fields como dict."""
    poller = client.begin_analyze_document(MODEL_ID, document=file_bytes)
    result = poller.result()

    if not result.documents:
        return {}

    doc = result.documents[0]
    fields = {}

    for name, field in doc.fields.items():
        value = field.value if field.value is not None else field.content
        fields[name] = value

    return fields


def build_notepad_text(fields: dict) -> str:
    """
    Monta o 'bloco de notas'.
    Primeiro, genérico: nome_do_campo: valor.
    Depois você pode customizar com a ordem exata do S360.
    """
    linhas = []
    for name, value in fields.items():
        linhas.append(f"{name}: {value}")
    return "\n".join(linhas)


# =========================
#  UI STREAMLIT
# =========================
st.title("OCR – Cartão de Óleo (Modelo: oil-card)")

st.write(
    "Faça upload da imagem ou PDF do cartão.\n"
    "O app usa o modelo **oil-card** do Azure Document Intelligence "
    "e gera um texto para você copiar e colar na tela do S360."
)

uploaded_file = st.file_uploader(
    "Envie o arquivo do cartão",
    type=["jpg", "jpeg", "png", "pdf"],
)

if uploaded_file is not None:
    # Preview se for imagem
    if uploaded_file.type.startswith("image/"):
        st.image(uploaded_file, caption="Imagem enviada")

    if st.button("Extrair dados com OCR"):
        with st.spinner("Processando com OCR..."):
            file_bytes = uploaded_file.getvalue()

            try:
                fields = extract_fields(file_bytes)
            except Exception as e:
                st.error(f"Erro ao chamar o modelo: {e}")
                st.stop()

        if not fields:
            st.warning("Nenhum campo retornado pelo modelo. Verifique o modelo ou o arquivo.")
        else:
            st.subheader("Campos retornados (JSON bruto)")
            st.json(fields)

            # Bloco de notas
            text_output = build_notepad_text(fields)

            st.subheader("Bloco de notas para copiar/colar no S360")
            st.text_area("Texto gerado", text_output, height=260)

            st.info("Dica: use Ctrl+A / Ctrl+C para copiar o texto acima.")
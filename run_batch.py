import sys, types, io, os, fitz, time
import pandas as pd

# Read Azure credentials from .streamlit/secrets.toml (gitignored — never hardcode
# keys here). Falls back to environment variables if the file is absent.
def _load_secrets():
    path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    try:
        import tomllib
        with open(path, "rb") as f:
            return tomllib.load(f)
    except Exception:
        return {
            "AZURE_DI_ENDPOINT": os.environ.get("AZURE_DI_ENDPOINT", ""),
            "AZURE_DI_KEY":      os.environ.get("AZURE_DI_KEY", ""),
            "AZURE_DI_MODEL_ID": os.environ.get("AZURE_DI_MODEL_ID", "oil-card3"),
        }

st = types.ModuleType('streamlit')
st.secrets = _load_secrets()
noop = lambda *a, **kw: None
for attr in [
    'cache_data','cache_resource','title','header','subheader','write',
    'info','warning','error','success','stop','markdown','divider',
    'expander','columns','sidebar','file_uploader','button','spinner',
    'dataframe','table','download_button','set_page_config','tabs',
    'progress','empty','metric','text','caption','code','json',
    'radio','selectbox','multiselect','checkbox','slider','number_input',
    'text_input','text_area','form','form_submit_button','balloons',
    'toast','status',
]:
    setattr(st, attr, noop)
st.cache_data    = lambda **kw: (lambda f: f)
st.cache_resource = lambda **kw: (lambda f: f)

class FakeCtx:
    def __enter__(self): return self
    def __exit__(self, *a): pass
    def __call__(self, *a, **kw): return self
    write = info = warning = error = success = markdown = noop

fake_ctx = FakeCtx()
for attr in ['expander', 'columns', 'sidebar', 'spinner', 'status', 'form', 'empty']:
    setattr(st, attr, fake_ctx)
st.session_state = {}
sys.modules['streamlit'] = st

sys.path.insert(0, r'C:\Users\abelh\Downloads\OCR_embedding\OCR_embedding')
import app

# ── Patch functions to trace compartment origin ──────────────────────────────
_orig_page_marks  = app._find_compartment_via_page_marks
_orig_pixel       = app._find_compartment_via_pixel_detection

_last_source = {"v": ""}

def _traced_page_marks(page, page_bytes=None):
    r = _orig_page_marks(page, page_bytes=page_bytes)
    if r:
        _last_source["v"] = f"page_marks={r}"
    return r

def _traced_pixel(pb, page):
    r, cat = _orig_pixel(pb, page)
    if r:
        _last_source["v"] = f"pixel={r}"
    return r, cat

app._find_compartment_via_page_marks  = _traced_page_marks
app._find_compartment_via_pixel_detection = _traced_pixel

# Trace _normalize_oil_description for debugging

_orig_fields_to_record = app.fields_to_record
def _traced_fields(fields, **kw):
    _last_source["v"] = ""
    # Show all compartment KV fields with their value and confidence
    import unicodedata, re
    def _n(s):
        s = str(s).lower().strip()
        s = unicodedata.normalize('NFD', s)
        s = ''.join(ch for ch in s if not unicodedata.combining(ch))
        s = re.sub(r'\s+', ' ', s)
        s = re.sub(r'[^a-z0-9 /?]', '', s)
        return s.strip()
    confs = kw.get("confidences") or {}
    comp_hits = []
    for k, v in (fields or {}).items():
        if v is None: continue
        canon = app.COMPARTMENT_OPTIONS.get(_n(k))
        if not canon: continue
        conf = confs.get(k, 1.0) or 1.0
        vstr = str(v).strip().lower()
        if vstr in ("selected","sim","si","yes","true","1","x"):
            comp_hits.append(f"{canon}:SEL@{conf:.2f}")
            if not _last_source["v"]:
                _last_source["v"] = f"azure_kv={canon}@{conf:.2f}"
    if comp_hits:
        _last_source["kv_detail"] = " | ".join(comp_hits)
    rec = _orig_fields_to_record(fields, **kw)
    return rec
app.fields_to_record = _traced_fields

# ── Process PDF ──────────────────────────────────────────────────────────────
# Ajuste estes caminhos para o PDF que quer processar em lote.
PDF_PATH  = os.environ.get("OCR_BATCH_PDF", r"entrada.pdf")
OUT_CSV   = r"resultado_batch.csv"
OUT_EXCEL = r"resultado_batch.xlsx"
MAX_PAGES = 20

doc   = fitz.open(PDF_PATH)
total = min(MAX_PAGES, len(doc))
print(f'Processando {total} de {len(doc)} paginas...\n')

records = []
for i in range(total):
    buf = io.BytesIO()
    sub = fitz.open()
    sub.insert_pdf(doc, from_page=i, to_page=i)
    sub.save(buf)
    sub.close()
    page_bytes = buf.getvalue()

    try:
        app._PM_DEBUG = False
        result = app.analyze_bytes(page_bytes)
        recs, _ = app.result_to_records(result, page_bytes=page_bytes)
        for r in recs:
            r['Pagina'] = i + 1
            records.append(r)
        comp    = recs[0].get('Ponto de Coleta / Compartimento', '') if recs else ''
        cat     = recs[0].get('Categoria', '') if recs else ''
        horim   = recs[0].get('Horímetro/Km/Período', '') if recs else ''
        hkm     = recs[0].get('Horas/Km do Fluído', '') if recs else ''
        vol     = recs[0].get('Volume adicionado', '') if recs else ''
        oleo    = recs[0].get('Fabricante/Modelo', '') if recs else ''
        visc    = recs[0].get('Viscosidade (Óleo)', '') if recs else ''
        source  = _last_source.get("v") or ('azure_kv' if comp else '')
        kv_det  = _last_source.get("kv_detail", "")
        flag    = ' *** MANCAL ***' if 'mancal' in comp.lower() else ''
        detail  = f'  [{kv_det}]' if kv_det else ''
        cat_info   = f'  cat={cat!r}' if cat else ''
        horas_info = f'  hkm={hkm!r}' if hkm else ''
        vol_info   = f'  vol={vol!r}' if vol else ''
        oil_info   = f'  oleo={oleo!r}' if oleo else ''
        visc_info  = f'  visc={visc!r}' if visc else ''
        print(f'  p{i+1:02d}: comp={comp!r:20s}{cat_info}  horim={horim!r:12s}{horas_info}{vol_info}{oil_info}{visc_info}')
        _last_source.clear()
    except Exception as e:
        print(f'  p{i+1:02d}: ERRO: {e}')
    time.sleep(0.3)

doc.close()
print(f'\nTotal: {len(records)} registros')

if records:
    cols = ['Pagina'] + app.COLETA_COLUMNS
    df = pd.DataFrame(records)
    df = df[[c for c in cols if c in df.columns]]
    df.to_csv(OUT_CSV, index=False, encoding='utf-8-sig')
    excel_bytes = app.build_excel_bytes(records)
    with open(OUT_EXCEL, 'wb') as f:
        f.write(excel_bytes)
    print(f'CSV  -> {OUT_CSV}')
    print(f'Excel-> {OUT_EXCEL}')
    print('\n--- Tabela ---')
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 200)
    pd.set_option('display.max_colwidth', 30)
    cols_show = ['Pagina','Ponto de Coleta / Compartimento','Data da Coleta',
                 'Chassi Série','Tag Frota','Horímetro/Km/Período','Horas/Km do Fluído',
                 'Volume adicionado','Fabricante/Modelo','Viscosidade (Óleo)']
    print(df[[c for c in cols_show if c in df.columns]].to_string(index=False))

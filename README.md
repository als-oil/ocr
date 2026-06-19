# OCR — Cartão de Óleo → Excel

App em Streamlit que recebe cartões/formulários de coleta de óleo (imagem ou PDF com várias páginas) e extrai os dados para uma planilha Excel formatada, usando um modelo customizado do **Azure Document Intelligence**.

## Funcionalidades

- Extração de cartões em PT/ES; páginas que não são formulário são filtradas automaticamente.
- Detecção de compartimento e "óleo trocado" por marca de tinta (cor), robusta a manchas no papel.
- Coluna **Categoria** (Móvel/Industrial) e separação de **Fabricante/Modelo** e **Viscosidade** em colunas distintas.
- Correção de OCR via catálogo de óleos (`oil_reference.csv`) e regras de plausibilidade por campo.
- **Confiabilidade por célula**: células pouco confiáveis ficam destacadas em laranja no Excel para revisão manual.

## Configuração

As credenciais do Azure ficam em `.streamlit/secrets.toml`, que **não é versionado** (ver `.gitignore`).

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edite secrets.toml com seu endpoint, chave e model id do Azure
```

```toml
AZURE_DI_ENDPOINT = "https://SEU-RECURSO.cognitiveservices.azure.com/"
AZURE_DI_KEY = "SUA_CHAVE_AQUI"
AZURE_DI_MODEL_ID = "oil-card3"
```

## Como rodar

```bash
python -m venv .venv
# Windows:  .\.venv\Scripts\Activate.ps1
# Linux/Mac: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

O app abre em `http://localhost:8501`. Envie um cartão (imagem) ou PDF e baixe o Excel.

## Estrutura

| Arquivo | Função |
|---|---|
| `app.py` | Aplicação Streamlit + pipeline de extração |
| `run_batch.py` | Processa um PDF em lote no terminal (debug/teste) |
| `oil_reference.csv` | Catálogo fabricante/modelo/viscosidade para correção de OCR |
| `requirements.txt` | Dependências Python |

## Segurança

Nunca faça commit de `.streamlit/secrets.toml` nem de chaves Azure em qualquer arquivo. Use sempre `secrets.toml.example` como modelo.

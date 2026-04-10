
# SRI UDESC CCT

Projeto Streamlit para consulta de temporalidade, inventário documental, eliminação documental e geração de capas/etiquetas de caixa.

## Base documental oficial
A pasta `data/reference/cdoc` contém os modelos oficiais usados pelo sistema:
- Anexo II - Atividades Fim
- Anexo V - Atividades Meio
- Anexo IV - Inventário
- Anexo I - Etiqueta/Capa de caixa
- TTD_UDESC_SEA_filtravel.xlsx como base consolidada para busca e preenchimento automático

## Como executar
```bash
python -m venv .venv
# Windows PowerShell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m streamlit run app.py
```

from __future__ import annotations

from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="SRI UDESC",
    page_icon="📁",
    layout="wide"
)

logo_path = Path(__file__).resolve().parent / "assets" / "logo_udesc.jpg"

col1, col2 = st.columns([1, 6])

with col1:
    if logo_path.exists():
        st.image(str(logo_path), width=120)

with col2:
    st.title("SRI – Sistema de Recuperação da Informação Arquivística")
    st.subheader("Arquivo Permanente | UDESC")

st.markdown("""
Bem-vindo ao Sistema de Recuperação da Informação Arquivística da UDESC.

Este sistema foi desenvolvido para apoiar a consulta da classificação documental,
dos prazos de guarda e da destinação final dos documentos produzidos e acumulados
no âmbito da Universidade do Estado de Santa Catarina.

## Objetivos do sistema
- facilitar a consulta à Tabela de Temporalidade de Documentos (TTD);
- apoiar a organização de arquivos correntes, intermediários e permanentes;
- auxiliar setores administrativos na identificação documental;
- contribuir para inventário, etiquetagem, acondicionamento e descarte documental.

## Módulos disponíveis
- **Consulta de Temporalidade**: permite pesquisar documentos e verificar classificação, prazos e destinação;
- **Inventário**: auxilia na elaboração de inventários documentais por setor/proveniência;
- **Etiquetas e Capas**: apoia a identificação de caixas, pastas e dossiês por setor/proveniência;
- **Descarte Documental**: fornece apoio para procedimentos de eliminação documental por setor/proveniência.

## Antes de iniciar a consulta
A consulta à temporalidade documental é organizada em dois grandes grupos:

### Atividade-meio
Refere-se às funções administrativas e de apoio da instituição, como:
- recursos humanos;
- compras;
- contratos;
- patrimônio;
- finanças;
- protocolo;
- gestão documental.

### Atividade-fim
Refere-se às funções ligadas à finalidade institucional da universidade, como:
- ensino;
- pesquisa;
- extensão;
- pós-graduação;
- registro acadêmico;
- atividades pedagógicas e finalísticas.

## Como utilizar o sistema
1. Acesse, no menu lateral, o módulo desejado.
2. No módulo **Consulta de Temporalidade**, selecione o tipo de atividade.
3. Utilize os filtros de classificação ou a pesquisa textual.
4. Consulte os prazos de guarda e a destinação final do documento.
5. Nos módulos operacionais, selecione sempre o **setor/proveniência** correspondente antes de gerar inventário, capas ou descarte documental.
6. Para saber mais, consulte o site da Coordenadoria de Documentação - CDOC (UDESC):
   https://www.udesc.br/proreitoria/proplan/cdoc

## Exemplo prático
- Um documento relacionado a **férias de servidor** tende a estar em **atividade-meio**.
- Um documento relacionado a **projeto de extensão** tende a estar em **atividade-fim**.

Desenvolvido por KP (CCT/UDESC) 2026

Use o menu lateral para navegar entre os módulos do sistema.
""")

st.info("Selecione um módulo no menu lateral para iniciar a navegação no sistema.")
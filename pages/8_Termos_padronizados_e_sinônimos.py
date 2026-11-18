import streamlit as st
import pandas as pd
from services.db import get_supabase_client, normalizar_tipo_atividade
# Importa a função que criamos acima
from services.ia_vocabulario import sugerir_equivalencias_ia

st.title("🤖 Moderação Inteligente de Vocabulário com IA")

supabase = get_supabase_client()

dados = (
    supabase
    .table("vocabulario_pendente")
    .select("*")
    .eq("aprovado", False)
    .execute()
)

df = pd.DataFrame(dados.data)

if df.empty:
    st.success("Nenhuma pendência de vocabulário encontrada.")
    st.stop()

st.info(f"{len(df)} termo(s) pendente(s) aguardando moderação.")

for _, row in df.iterrows():
    documento_faltante = row["documento_faltante"]
    
    st.markdown(f"### 📌 Termo buscado: `{documento_faltante}`")

    # Se a base ainda não tem termo semelhante cadastrado, podemos rodar a IA sob demanda para ajudar o revisor!
    termo_sugerido_ia = row.get("termo_semelhante")
    lista_sinonimos_ia = [documento_faltante]

    if not termo_sugerido_ia or termo_sugerido_ia == "":
        with st.spinner("🤖 Consultando IA arquivista para sugerir sinônimos e termo padronizado..."):
            dados_ia = sugerir_equivalencias_ia(documento_faltante)
            termo_sugerido_ia = dados_ia.get("termo_padronizado", documento_faltante)
            lista_sinonimos_ia = dados_ia.get("sinonimos", [documento_faltante])

    col1, col2 = st.columns(2)
    
    with col1:
        # O campo já nasce preenchido inteligentemente pela IA
        termo_padronizado = st.text_input(
            "Termo Padronizado Oficial",
            value=str(termo_sugerido_ia),
            key=f"termo_padronizado_{row['id']}",
        )

    with col2:
        assunto = st.text_input(
            "Assunto Técnico",
            value=documento_faltante,
            key=f"assunto_{row['id']}"
        )

    # Transforma a lista de sinônimos da IA em texto com quebra de linha
    texto_sinonimos_inicial = "\n".join(dict.fromkeys(lista_sinonimos_ia))

    sinonimos = st.text_area(
        "Sinônimos e Siglas sugeridos (um por linha)",
        value=texto_sinonimos_inicial,
        key=f"sinonimos_{row['id']}",
        help="A IA gerou essas opções com base no contexto arquivístico. Você pode editar livremente antes de aprovar.",
    )

    observacao = st.text_area(
        "Observação do revisor",
        key=f"obs_{row['id']}"
    )

    if st.button(f"✅ Aprovar e Cadastrar Termo #{row['id']}", key=f"aprovar_{row['id']}", type="primary"):
        termo_preferido = termo_padronizado.strip() or documento_faltante
        
        termos_busca = [
            termo.strip()
            for termo in sinonimos.splitlines()
            if termo.strip()
        ]

        if termo_preferido not in termos_busca:
            termos_busca.insert(0, termo_preferido)

        tipo_atividade = normalizar_tipo_atividade(row.get("tipo_atividade") or "fim")

        # Grava no vocabulário controlado do Supabase
        for termo_busca in dict.fromkeys(termos_busca):
            supabase.table("vocabulario_controlado").insert(
                {
                    "tipo_atividade": tipo_atividade,
                    "termo_encontrado": termo_busca,
                    "termo_padronizado": termo_preferido,
                    "atividade": "",
                    "assunto": assunto,
                    "codigo_classificacao": "",
                }
            ).execute()

        # Prepara o dicionário de atualização de forma segura
        dados_atualizacao = {
            "aprovado": True, 
            "revisado": True
        }

        # Verifica se a observação existe na tela e foi preenchida
        observacao_digitada = locals().get("observacao", "").strip()
        if observacao_digitada:
            dados_atualizacao["observacao"] = observacao_digitada

        # Executa a chamada ao Supabase com o dicionário dinâmico
        supabase.table("vocabulario_pendente").update(dados_atualizacao).eq("id", row["id"]).execute()

        st.success("Termo e sinônimos cadastrados com sucesso!")
        st.rerun()

    st.divider()    
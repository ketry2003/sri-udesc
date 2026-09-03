import streamlit as st
import pandas as pd

from services.db import get_supabase_client

st.title("Vocabulário")

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
    st.success("Nenhuma pendência encontrada.")
    st.stop()

st.info(
    f"{len(df)} termo(s) pendente(s) de revisão."
)

with st.expander("Visualizar pendências"):
    st.dataframe(
        df,
        use_container_width=True
    )

st.divider()

selecionados = []

aprovar_selecionados = st.button(
    "Aprovar termos selecionados",
    type="primary",
    key="aprovar_termos_selecionados",
)

for _, row in df.iterrows():

    st.subheader(row["documento_faltante"])

    selecionado = st.checkbox(
        "Selecionar para aprovação em lote",
        key=f"selecionar_{row['id']}",
    )

    if selecionado:
        selecionados.append(row["id"])

    st.write(
        f"Score de similaridade: "
        f"{row.get('score') or '-'}"
    )

    st.write(
        f"Termo semelhante encontrado: "
        f"{row.get('termo_semelhante') or 'Nenhuma sugestão automática'}"
    )

    termo_padronizado = st.text_input(
        "Termo preferido",
        value=str(row.get("termo_semelhante", "") or row["documento_faltante"]),
        key=f"termo_padronizado_{row['id']}",
    )

    assunto = st.text_input(
        "Assunto",
        value=row["documento_faltante"],
        key=f"assunto_{row['id']}"
    )

    observacao = st.text_area(
        "Observação do revisor",
        key=f"obs_{row['id']}"
    )

    col1, col2 = st.columns(2)

    with col1:

        aprovar_individualmente = st.button(
            f"Aprovar #{row['id']}",
            key=f"aprovar_{row['id']}"
        )

        if aprovar_individualmente or (
            aprovar_selecionados and row["id"] in selecionados
        ):

            resultado = (
                supabase
                .table("vocabulario_controlado")
                .insert(
                    {
                        "tipo_atividade":
                            row.get("tipo_atividade") or "fim",

                        "termo_encontrado":
                            row["documento_faltante"],

                        "termo_padronizado":
                            termo_padronizado.strip()
                            or row["documento_faltante"],

                        "atividade": "",

                        "assunto":
                            assunto,

                        "codigo_classificacao": ""
                    }
                )
                .execute()
            )

            supabase.table(
                "vocabulario_pendente"
            ).update(
                {
                    "aprovado": True,
                    "revisado": True
                }
            ).eq(
                "id",
                row["id"]
            ).execute()

            st.success(
                "Termo aprovado com sucesso."
                if aprovar_individualmente
                else "Termo incluído na aprovação em lote."
            )

            if aprovar_individualmente:
                st.rerun()

    with col2:

        if st.button(
            f"❌ Rejeitar #{row['id']}",
            key=f"rejeitar_{row['id']}"
        ):

            supabase.table(
                "vocabulario_pendente"
            ).update(
                {
                    "revisado": True,
                    "aprovado": False,
                    "observacao": observacao
                }
            ).eq(
                "id",
                row["id"]
            ).execute()

            st.success(
                "Termo rejeitado."
            )

            st.rerun()

    st.divider()

if aprovar_selecionados:
    if not selecionados:
        st.warning("Selecione pelo menos um termo para aprovação em lote.")
    else:
        st.success(f"{len(selecionados)} termo(s) aprovado(s) com sucesso.")
        st.rerun()

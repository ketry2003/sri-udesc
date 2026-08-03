import streamlit as st
import pandas as pd

from services.db import get_supabase_client

st.title("Auditoria do Vocabulário")

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

for _, row in df.iterrows():

    st.subheader(row["documento_faltante"])

    st.write(
        f"Score de similaridade: "
        f"{row.get('score', '-')}"
    )

    st.write(
        f"Termo semelhante encontrado: "
        f"{row.get('termo_semelhante', '-')}"
    )

    tipo_atividade = st.selectbox(
        "Tipo de atividade",
        ["fim", "meio"],
        key=f"tipo_{row['id']}"
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

        if st.button(
            f"Aprovar #{row['id']}",
            key=f"aprovar_{row['id']}"
        ):

            resultado = (
                supabase
                .table("vocabulario_controlado")
                .insert(
                    {
                        "tipo_atividade":
                            tipo_atividade,

                        "termo_encontrado":
                            row["documento_faltante"],

                        "termo_padronizado":
                            row["documento_faltante"],

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
            )

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
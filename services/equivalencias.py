import pandas as pd

from services.db import (
    buscar_equivalencia_historica,
    salvar_equivalencia_historica,
    carregar_vocabulario
)

from services.similaridade import (
    encontrar_melhor_termo
)


def buscar_equivalencia(termo):

    equivalente = buscar_equivalencia_historica(termo)

    if equivalente:
        return {
            "termo": equivalente,
            "score": 100
        }

    df_fim = carregar_vocabulario("Atividade-fim")
    df_meio = carregar_vocabulario("Atividade-meio")

    df_vocab = pd.concat(
        [df_fim, df_meio],
        ignore_index=True
    )

    if df_vocab.empty:
        return None

    colunas = [
        "assunto",
        "termo_padronizado",
        "termo_encontrado",
        "atividade",
        "sinonimo"
    ]

    termos = []

    for coluna in colunas:
        if coluna in df_vocab.columns:
            termos.extend(
                df_vocab[coluna]
                .dropna()
                .astype(str)
                .tolist()
            )

    sugestao = encontrar_melhor_termo(
        termo,
        termos,
        score_minimo=70
    )

    if sugestao:
        return sugestao

    return None


def salvar_equivalencia(
    termo_historico,
    termo_oficial,
    observacao=""
):
    return salvar_equivalencia_historica(
        termo_historico,
        termo_oficial,
        observacao
    )

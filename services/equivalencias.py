from services.db import (
    buscar_equivalencia_historica,
    salvar_equivalencia_historica,
    carregar_vocabulario
)

from services.similaridade import (
    encontrar_melhor_termo
)

def buscar_equivalencia(termo):

    equivalente = buscar_equivalencia_historica(
        termo
    )

    if equivalente:
        return equivalente

    df_vocab = carregar_vocabulario(
        "Atividade-fim"
    )

    termos = (
        df_vocab["assunto"]
        .dropna()
        .astype(str)
        .tolist()
    )

    sugestao = encontrar_melhor_termo(
        termo,
        termos,
        score_minimo=85
    )

    if sugestao:
        return sugestao["termo"]

    return None
from rapidfuzz import fuzz


def encontrar_melhor_termo(
    termo_digitado,
    lista_termos,
    score_minimo=85
):
    melhor_termo = None
    melhor_score = 0

    for termo in lista_termos:
        score = fuzz.token_set_ratio(
            str(termo_digitado),
            str(termo)
        )

        if score > melhor_score:
            melhor_score = score
            melhor_termo = termo

    if melhor_score >= score_minimo:
        return {
            "termo": melhor_termo,
            "score": melhor_score,
        }

    return None

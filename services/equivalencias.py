from services.db import (
    buscar_equivalencia_historica,
    salvar_equivalencia_historica,
)


def buscar_equivalencia(
    termo
):
    return buscar_equivalencia_historica(
        termo
    )


def salvar_equivalencia(
    termo_historico,
    termo_oficial,
    observacao=""
):
    return salvar_equivalencia_historica(
        termo_historico,
        termo_oficial,
        observacao,
    )
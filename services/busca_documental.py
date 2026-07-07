from rapidfuzz import fuzz
import pandas as pd


def buscar_documentos(df, consulta, limite=20):

    consulta = str(consulta).strip()

    if not consulta:
        return pd.DataFrame()

    resultados = []

    for _, row in df.iterrows():

        texto = " ".join([
            str(row.get("termo_preferido_oficial", "")),
            str(row.get("termos_populares_sugeridos", "")),
            str(row.get("pergunta_guia_usuario", "")),
            str(row.get("assunto_tecnico", "")),
            str(row.get("funcao", "")),
            str(row.get("subfuncao", "")),
            str(row.get("atividade", "")),
            str(row.get("codigo_classificacao", "")),
            str(row.get("observacao", "")),
            str(row.get("texto_busca_sistema", "")),
        ])

        score = fuzz.WRatio(
            consulta.lower(),
            texto.lower()
        )

        if score >= 55:

            linha = row.copy()

            linha["score"] = score

            resultados.append(linha)

    if not resultados:
        return pd.DataFrame()

    resultado = pd.DataFrame(resultados)

    resultado = resultado.sort_values(
        "score",
        ascending=False
    )

    return resultado.head(limite)
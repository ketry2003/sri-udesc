from rapidfuzz import fuzz

def buscar_documentos(df, consulta):

    resultados = []

    consulta = str(consulta)

    for _, row in df.iterrows():

        texto = " ".join([
            str(row.get("termo_preferido_oficial", "")),
            str(row.get("termos_populares_sugeridos", "")),
            str(row.get("assunto_tecnico", "")),
        ])

        score = fuzz.WRatio(
            consulta,
            texto
        )

        if score >= 60:

            resultados.append(
                (
                    score,
                    row
                )
            )

    resultados.sort(
        reverse=True,
        key=lambda x: x[0]
    )

    return resultados[:20]
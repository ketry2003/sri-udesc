documentos_antigos = set(
    ttd[ttd["arquivo"] == "Atividade-fim UDESC"]["documento"]
)

documentos_novos = set(
    fim["documento"]
)

novos = documentos_novos - documentos_antigos
removidos = documentos_antigos - documentos_novos

print(f"Novos: {len(novos)}")
print(f"Removidos: {len(removidos)}")
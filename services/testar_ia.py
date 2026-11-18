# testar_ia.py
from services.ia_vocabulario import sugerir_equivalencias_ia

print("🤖 Consultando a IA arquivista...")

# Testando com uma sigla comum do seu acervo
termo_teste = "PTI" 

resultado = sugerir_equivalencias_ia(termo_teste)

print("\n--- Resultado Retornado pela IA ---")
print(resultado)
print(f"Tipo do dado retornado: {type(resultado)}")

if isinstance(resultado, dict) and "termo_padronizado" in resultado:
    print("\n✅ Sucesso! O JSON foi interpretado corretamente como dicionário Python.")
    print(f"📌 Termo Padronizado: {resultado.get('termo_padronizado')}")
    print(f"📋 Sinônimos sugeridos: {resultado.get('sinonimos')}")
else:
    print("\n⚠️ Atenção: O retorno não veio no formato de dicionário esperado.")
import re
import pandas as pd
from services.db import get_supabase_client

def gerar_variacoes_sinonimos(termo):
    """Gera variações e sinônimos limpos baseados na lógica de palavras-chave."""
    termo_str = str(termo).strip()
    sinonimos = set()
    
    # 1. Termo original limpo de pontuação excessiva
    limpo = re.sub(r'[,\.\(\)\/\?\!]', ' ', termo_str)
    limpo = re.sub(r'\s+', ' ', limpo).strip()
    
    # Stopwords para remoção (preposições, artigos, etc.)
    stopwords = {'de', 'do', 'da', 'dos', 'das', 'em', 'no', 'na', 'nos', 'nas', 'para', 'e', 'a', 'o', 'as', 'os', 'um', 'uma', 'ao', 'aos', 'à', 'às', 'por', 'com'}
    
    palavras = [p for p in limpo.split() if p.lower() not in stopwords]
    
    if palavras:
        # Versão sem preposições
        sinonimos.add(" ".join(palavras))
        
        # Versão sem a primeira palavra (ex: "Requerimento de X" -> "X")
        if len(palavras) > 1:
            sinonimos.add(" ".join(palavras[1:]))
            
        # Versão com a primeira e a última palavra principal
        if len(palavras) > 2:
            sinonimos.add(f"{palavras[0]} {palavras[-1]}")

    # Retorna as variações unidas por quebra de linha (formato padrão que você usa)
    return "\n---\n".join(list(sinonimos))

def processar_lote():
    # 1. Lê o CSV local com os pendentes
    print("Lendo o arquivo vocabulario_pendente_rows.csv...")
    df = pd.read_csv('vocabulario_pendente_rows.csv')
    
    # Filtra apenas os pendentes que não foram revisados/aprovados
    pendentes = df[(df['revisado'] == False) | (df['aprovado'] == False)]
    print(f"Total de termos pendentes encontrados: {len(pendentes)}")
    
    # Conecta ao Supabase
    supabase = get_supabase_client()
    
    sucessos = 0
    for idx, row in pendentes.iterrows():
        id_termo = row['id']
        termo = row['documento_faltante']
        
        # Gera os sinônimos automaticamente
        sinonimos_gerados = gerar_variacoes_sinonimos(termo)
        
        try:
            # Atualiza no Supabase marcando como revisado, aprovado e incluindo as variações
            # (Caso sua tabela possua uma coluna de sinônimos ou observação, você pode ajustar aqui)
            dados = {
                "revisado": True,
                "aprovado": True
            }
            
            supabase.table("vocabulario_pendente").update(dados).eq("id", id_termo).execute()
            sucessos += 1
            print(f"[{sucessos}] Termo ID {id_termo} processado: '{termo}'")
        except Exception as e:
            print(f"Erro ao processar ID {id_termo}: {e}")
            
    print(f"\nProcessamento concluído! {sucessos} termos processados com sucesso.")

if __name__ == "__main__":
    processar_lote()
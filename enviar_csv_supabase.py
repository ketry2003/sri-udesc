import pandas as pd
import numpy as np
from services.db import get_supabase_client

def enviar_csv():
    # 1. Carrega o arquivo CSV local
    print("Lendo o arquivo vocabulario_pendente_rows.csv...")
    df = pd.read_csv('vocabulario_pendente_rows.csv')
    
    # Remove a coluna 'id' para deixar o Supabase gerar automaticamente via auto-incremento
    if 'id' in df.columns:
        df = df.drop(columns=['id'])
    
    # Substitui valores NaN/vazios por None (compatíveis com JSON/Supabase)
    df = df.where(pd.notnull(df), None)
    
    # Converte o DataFrame para lista de dicionários
    registros = df.to_dict(orient='records')
    total = len(registros)
    print(f"Total de registros a enviar: {total}")
    
    # 2. Conecta ao Supabase
    supabase = get_supabase_client()
    
    # 3. Envia em blocos de 50 para evitar sobrecarga na API
    tamanho_lote = 50
    erros = 0
    for i in range(0, total, tamanho_lote):
        lote = registros[i:i + tamanho_lote]
        try:
            # Utiliza insert para novos registros sem ID fixo
            response = supabase.table("vocabulario_pendente").insert(lote).execute()
            print(f"Enviado lote {i} até {min(i + tamanho_lote, total)} de {total} com sucesso.")
        except Exception as e:
            erros += 1
            print(f"Erro ao enviar o lote {i}: {e}")
            
    if erros == 0:
        print("\nProcesso de envio para o Supabase concluído com sucesso!")
    else:
        print(f"\nProcesso concluído com {erros} erros.")

if __name__ == "__main__":
    enviar_csv()
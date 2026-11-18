import os
import json
from google import genai

def sugerir_equivalencias_ia(termo_faltante: str) -> dict:
    """
    Usa a API gratuita do Google Gemini (gemini-2.5-flash) para atuar 
    como um arquivista e sugerir um termo padronizado e sinônimos.
    """
    api_key = os.getenv("GEMINI_API_KEY") 
    if not api_key:
        try:
            import streamlit as st
            api_key = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            api_key = ""

    if not api_key:
        return {"termo_padronizado": termo_faltante, "sinonimos": [termo_faltante]}

    try:
        client = genai.Client(api_key=api_key)

        prompt = f"""
        Você é um arquivista especialista na Tabela de Temporalidade e Destinação (TTD) de órgãos públicos e universidades (como a UDESC).
        O usuário pesquisou por um termo que não foi encontrado no vocabulário controlado: "{termo_faltante}".
        
        Retorne estritamente em formato JSON contendo:
        1. "termo_padronizado": Uma sugestão de nome formal e oficial para este documento (ex: "Plano de Trabalho Individual").
        2. "sinonimos": Uma lista com 3 a 5 variações informais, siglas (como PTI) ou formas de busca que os servidores costumam usar.
        
        Retorne APENAS o JSON válido, sem blocos markdown (como ```json) ou textos adicionais.
        """

        response = client.models.generate_content(
            model="gemini-3.8-flash",  # Modelo rápido, inteligente e gratuito do Google
            contents=prompt,
        )
        
        # Remove eventuais marcações de bloco de código da resposta
        texto_limpo = response.text.replace("```json", "").replace("```", "").strip()
        resultado = json.loads(texto_limpo)
        return resultado
        
    except Exception as e:
        print(f"Erro ao consultar Gemini: {e}")
        return {"termo_padronizado": termo_faltante, "sinonimos": [termo_faltante]}
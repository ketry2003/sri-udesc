from __future__ import annotations

import os
from datetime import datetime, timezone

try:
    import streamlit as st
except Exception:
    st = None

from supabase import create_client


def get_supabase_client():

    if st is not None:
        url = st.secrets.get("SUPABASE_URL", "")
        key = st.secrets.get("SUPABASE_KEY", "")
    else:
        url = os.getenv("SUPABASE_URL", "")
        key = os.getenv("SUPABASE_KEY", "")

    if not url or not key:
        raise RuntimeError(
            "SUPABASE_URL e SUPABASE_KEY não configurados."
        )

    try:
        return create_client(url, key)
    except Exception as exc:
        raise RuntimeError(f"Falha ao conectar ao Supabase: {exc}") from exc


def init_db():
    if st is not None:
        st.sidebar.success("Banco ativo: Supabase API")


def _response_data(response):
    return response.data or []


def insert_inventory_item(payload: dict) -> int:
    supabase = get_supabase_client()

    data = {
        "setor": payload.get("setor", ""),
        "tipo_documental": payload.get("tipo_documental", ""),
        "natureza_documental": payload.get("natureza_documental", ""),
        "grupo": payload.get("grupo", ""),
        "subgrupo": payload.get("subgrupo", ""),
        "serie": payload.get("serie", ""),
        "subserie": payload.get("subserie", ""),
        "dossie_processo": payload.get("dossie_processo", ""),
        "item_documental": payload.get("item_documental", ""),
        "codigo_classificacao": payload.get("codigo_classificacao", ""),
        "prazo_corrente": payload.get("prazo_corrente", ""),
        "prazo_intermediario": payload.get("prazo_intermediario", ""),
        "destinacao_final": payload.get("destinacao_final", ""),
        "datas_limite": payload.get("datas_limite", ""),
        "quantidade": int(payload.get("quantidade", 1) or 1),
        "caixa": str(payload.get("caixa", "")).strip(),
        "observacoes": payload.get("observacoes", ""),
    }

    response = (
        supabase
        .table("inventory_items")
        .insert(data)
        .execute()
    )

    rows = _response_data(response)

    if not rows:
        raise RuntimeError("Erro ao inserir item no Supabase.")

    return int(rows[0]["id"])


def find_duplicate_inventory_item(payload: dict):
    supabase = get_supabase_client()

    query = (
        supabase
        .table("inventory_items")
        .select("*")
        .eq("setor", payload.get("setor", ""))
        .eq("tipo_documental", payload.get("tipo_documental", ""))
        .eq("item_documental", payload.get("item_documental", ""))
        .eq("codigo_classificacao", payload.get("codigo_classificacao", ""))
        .eq("datas_limite", payload.get("datas_limite", ""))
        .limit(1)
        .execute()
    )

    rows = _response_data(query)
    return rows[0] if rows else None


def list_inventory_items():
    supabase = get_supabase_client()

    response = (
        supabase
        .table("inventory_items")
        .select("*")
        .order("id", desc=True)
        .execute()
    )

    return _response_data(response)


def list_setores_inventory():
    rows = list_inventory_items()

    setores = sorted(
        {
            str(row.get("setor", "")).strip()
            for row in rows
            if str(row.get("setor", "")).strip()
        }
    )

    return setores


def list_inventory_items_by_setor(setor):
    supabase = get_supabase_client()

    response = (
        supabase
        .table("inventory_items")
        .select("*")
        .eq("setor", setor)
        .order("id", desc=True)
        .execute()
    )

    return _response_data(response)


def registrar_emprestimo(payload: dict):
    supabase = get_supabase_client()

    inventory_id = payload.get("inventory_id")
    if inventory_id is not None and str(inventory_id).strip() == "":
        inventory_id = None

    usuario_nome = str(payload.get("usuario_nome", "")).strip()
    usuario_id_udesc = str(payload.get("usuario_id_udesc", "")).strip()
    setor_usuario = str(payload.get("setor_usuario", "")).strip()

    if not usuario_nome:
        raise ValueError("Nome do usuário é obrigatório.")

    if not setor_usuario:
        raise ValueError("Setor de trabalho do usuário é obrigatório.")

    item = None
    if inventory_id is not None:
        item = next(
            (
                item
                for item in list_inventory_items()
                if int(item.get("id")) == int(inventory_id)
            ),
            None,
        )

        if item is None:
            raise ValueError("Item do inventário não encontrado.")

        ja_emprestado = (
            supabase
            .table("document_loans")
            .select("id")
            .eq("inventory_id", int(inventory_id))
            .eq("status", "Emprestado")
            .limit(1)
            .execute()
        )

        if getattr(ja_emprestado, "data", None):
            raise RuntimeError(
                "Este documento já está emprestado e não pode ser emprestado novamente."
            )

    documento = str(
        payload.get("documento")
        or (item.get("item_documental") if item else "")
        or (item.get("tipo_documental") if item else "")
        or ""
    ).strip()

    if not documento:
        raise ValueError("Informe o documento para registrar o empréstimo.")

    setor = str(
        payload.get("setor") or (item.get("setor") if item else "") or ""
    ).strip()

    if not setor:
        raise ValueError("Setor é obrigatório.")

    agora = datetime.now(timezone.utc).isoformat()

    data = {
        "inventory_id": int(inventory_id) if inventory_id is not None else None,
        "usuario_nome": usuario_nome,
        "usuario_id_udesc": usuario_id_udesc,
        "setor_usuario": setor_usuario,
        "documento": documento,
        "codigo_classificacao": str(
            payload.get("codigo_classificacao")
            or (item.get("codigo_classificacao") if item else "")
            or ""
        ).strip(),
        "caixa": str(
            payload.get("caixa") or (item.get("caixa") if item else "") or ""
        ).strip(),
        "setor": setor,
        "data_emprestimo": payload.get("data_emprestimo") or agora,
        "data_devolucao": None,
        "status": "Emprestado",
        "observacoes": str(payload.get("observacoes", "") or "").strip(),
        "criado_em": payload.get("criado_em") or agora,
    }

    response = (
        supabase
        .table("document_loans")
        .insert(data)
        .execute()
    )

    rows = _response_data(response)

    if not rows:
        raise RuntimeError("Erro ao registrar empréstimo no Supabase.")

    return rows[0]


def listar_emprestimos():
    supabase = get_supabase_client()

    response = (
        supabase
        .table("document_loans")
        .select("*")
        .order("criado_em", desc=True)
        .execute()
    )

    return _response_data(response)


def listar_documentos_emprestados():
    supabase = get_supabase_client()

    response = (
        supabase
        .table("document_loans")
        .select("*")
        .eq("status", "Emprestado")
        .order("data_emprestimo", desc=True)
        .execute()
    )

    return _response_data(response)


def registrar_devolucao(loan_id, observacoes=""):
    supabase = get_supabase_client()

    if loan_id is None:
        raise ValueError("ID do empréstimo é obrigatório.")

    emprestimo = buscar_emprestimo_por_id(loan_id)
    if not emprestimo:
        raise ValueError("Empréstimo não encontrado.")

    agora = datetime.now(timezone.utc).isoformat()

    response = (
        supabase
        .table("document_loans")
        .update(
            {
                "status": "Devolvido",
                "data_devolucao": agora,
                "observacoes": (
                    str(emprestimo.get("observacoes", "") or "")
                    + (" | " if observacoes else "")
                    + str(observacoes or "")
                ).strip(),
            }
        )
        .eq("id", int(loan_id))
        .execute()
    )

    rows = _response_data(response)

    if not rows:
        raise RuntimeError("Erro ao registrar devolução no Supabase.")

    return rows[0]


def buscar_emprestimo_por_id(loan_id):
    supabase = get_supabase_client()

    if loan_id is None:
        return None

    response = (
        supabase
        .table("document_loans")
        .select("*")
        .eq("id", int(loan_id))
        .limit(1)
        .execute()
    )

    rows = _response_data(response)
    return rows[0] if rows else None


def get_next_caixa_by_setor(setor):
    if not setor or not str(setor).strip():
        return "001"

    itens = list_inventory_items_by_setor(setor)

    numeros = []

    for item in itens:
        caixa = str(item.get("caixa", "")).strip()

        if caixa.isdigit():
            numeros.append(int(caixa))

    proximo = max(numeros, default=0) + 1

    return str(proximo).zfill(3)


def delete_inventory_item(item_id):
    supabase = get_supabase_client()

    supabase.table("validation_records").delete().eq(
        "inventory_id",
        item_id
    ).execute()

    response = (
        supabase
        .table("inventory_items")
        .delete()
        .eq("id", item_id)
        .execute()
    )

    return 1 if response and getattr(response, "data", None) else 0


def delete_inventory_items(item_ids):
    if not item_ids:
        return 0

    deletados = 0
    for item_id in item_ids:
        deletados += delete_inventory_item(item_id)

    return deletados


def delete_inventory_items_by_setor(setor):
    itens = list_inventory_items_by_setor(setor)

    ids = [item["id"] for item in itens]

    return delete_inventory_items(ids)


def save_validation_record(
    inventory_id,
    situacao_prazo,
    pode_eliminar,
    justificativa,
):
    supabase = get_supabase_client()

    data = {
        "inventory_id": inventory_id,
        "situacao_prazo": situacao_prazo,
        "pode_eliminar": bool(pode_eliminar),
        "justificativa": justificativa,
    }

    supabase.table("validation_records").insert(
        data
    ).execute()


def list_elimination_candidates():
    supabase = get_supabase_client()

    response = (
        supabase
        .table("validation_records")
        .select("*, inventory_items(*)")
        .eq("pode_eliminar", True)
        .execute()
    )

    return _response_data(response)


def update_inventory_item(item_id, payload):
    supabase = get_supabase_client()

    data = {
        "setor": payload.get("setor", ""),
        "tipo_documental": payload.get("tipo_documental", ""),
        "natureza_documental": payload.get("natureza_documental", ""),
        "grupo": payload.get("grupo", ""),
        "subgrupo": payload.get("subgrupo", ""),
        "serie": payload.get("serie", ""),
        "subserie": payload.get("subserie", ""),
        "dossie_processo": payload.get("dossie_processo", ""),
        "item_documental": payload.get("item_documental", ""),
        "codigo_classificacao": payload.get("codigo_classificacao", ""),
        "prazo_corrente": payload.get("prazo_corrente", ""),
        "prazo_intermediario": payload.get("prazo_intermediario", ""),
        "destinacao_final": payload.get("destinacao_final", ""),
        "datas_limite": payload.get("datas_limite", ""),
        "quantidade": payload.get("quantidade", 1),
        "caixa": payload.get("caixa", ""),
        "observacoes": payload.get("observacoes", ""),
    }

    response = (
        supabase
        .table("inventory_items")
        .update(data)
        .eq("id", item_id)
        .execute()
    )

    rows = response.data if response is not None and getattr(response, "data", None) else []
    return len(rows)

def salvar_equivalencia_historica(
    termo_historico,
    termo_oficial,
    observacao=""
):
    supabase = get_supabase_client()

    data = {
        "termo_historico": termo_historico,
        "termo_oficial": termo_oficial,
        "observacao": observacao,
    }

    response = (
        supabase
        .table("equivalencias_historicas")
        .upsert(
            data,
            on_conflict="termo_historico"
        )
        .execute()
    )

    return _response_data(response)


def buscar_equivalencia_historica(
    termo_historico
):
    supabase = get_supabase_client()

    termo = str(termo_historico or "").strip()
    if not termo:
        return None

    response = (
        supabase
        .table("equivalencias_historicas")
        .select("*")
        .ilike("termo_historico", termo)
        .limit(1)
        .execute()
    )

    rows = _response_data(response)

    for row in rows:
        if str(row.get("termo_historico", "")).strip().upper() == termo.upper():
            return row.get("termo_oficial")

    return None

import pandas as pd


def normalizar_tipo_atividade(valor):
    if valor is None:
        return ""

    valor_texto = str(valor).strip()
    if not valor_texto:
        return ""

    valor_normalizado = (
        valor_texto
        .lower()
        .replace("_", " ")
        .replace("-", " ")
        .strip()
    )
    valor_normalizado = " ".join(valor_normalizado.split())

    if valor_normalizado in {"meio", "atividade meio"}:
        return "Atividade-meio"

    if valor_normalizado in {"fim", "atividade fim"}:
        return "Atividade-fim"

    return valor_texto

def search_inventory_global(termo_geral="", termo_detalhado=""):
    """
    Realiza busca no inventário:
    - termo_geral: busca em qualquer campo principal (setor, tipo documental, código, observações).
    - termo_detalhado: busca focada em campos específicos (observacoes).
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    query = "SELECT * FROM inventory_items WHERE 1=1"
    params = []
    
    if termo_geral.strip():
        like_geral = f"%{termo_geral.strip()}%"
        query += " AND (setor LIKE ? OR tipo_documental LIKE ? OR codigo_classificacao LIKE ? OR observacoes LIKE ?)"
        params.extend([like_geral, like_geral, like_geral, like_geral])
        
    if termo_detalhado.strip():
        like_detalhado = f"%{termo_detalhado.strip()}%"
        query += " AND observacoes LIKE ?"
        params.append(like_detalhado)
        
    query += " ORDER BY id DESC"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return rows

def carregar_vocabulario(tipo):
    supabase = get_supabase_client()
    tipo_padronizado = normalizar_tipo_atividade(tipo)

    response = (
        supabase
        .table("vocabulario_controlado")
        .select("*")
        .eq("tipo_atividade", tipo_padronizado)
        .execute()
    )

    data = response.data if response is not None and getattr(response, "data", None) is not None else []
    return pd.DataFrame(data)

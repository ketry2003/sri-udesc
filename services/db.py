from __future__ import annotations

import os

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

    return create_client(url, key)


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


def get_next_caixa_by_setor(setor):
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

    supabase.table("inventory_items").delete().eq(
        "id",
        item_id
    ).execute()


def delete_inventory_items(item_ids):
    if not item_ids:
        return 0

    for item_id in item_ids:
        delete_inventory_item(item_id)

    return len(item_ids)


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

    response = (
        supabase
        .table("equivalencias_historicas")
        .select("*")
        .ilike(
            "termo_historico",
            termo_historico
        )
        .limit(1)
        .execute()
    )

    rows = _response_data(response)

    if not rows:
        return None

    return rows[0]["termo_oficial"]

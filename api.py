"""
api.py — Requisições à API SGG (endpoint /funcionario/).

Busca todos os funcionários ativos de uma empresa identificada pelo CNPJ,
lidando automaticamente com a paginação da API (máximo 100 por página).

Referência da API: https://app.sgg.net.br/api/v3/doc/
"""

import json
import time
from typing import Any

import requests

from config import get_headers, get_url_funcionarios


class APIError(Exception):
    """Exceção customizada para erros de API."""
    pass


def _fazer_requisicao(url: str, payload: dict) -> Any:
    """
    Faz um GET autenticado com corpo JSON (padrão da API SGG).

    A API SGG aceita dados no corpo do GET (como data-raw) ou como
    query params. Aqui usamos o corpo JSON.
    """
    try:
        response = requests.get(
            url,
            headers=get_headers(),
            data=json.dumps(payload),
            timeout=60,
        )
        response.raise_for_status()
        return response.json()

    except requests.exceptions.ConnectionError:
        raise APIError(
            "Não foi possível conectar à API do SGG.\n"
            "Verifique sua conexão com a internet."
        )
    except requests.exceptions.Timeout:
        raise APIError("A requisição excedeu o tempo limite (60s).")
    except requests.exceptions.HTTPError as e:
        status = e.response.status_code if e.response is not None else "?"
        if status == 401:
            raise APIError(
                "Erro de autenticação (401).\n"
                "Verifique se a chave de API no .env está correta."
            )
        if status == 429:
            raise APIError(
                "Limite de requisições excedido (429).\n"
                "Aguarde um minuto e tente novamente."
            )
        body = ""
        try:
            body = e.response.json().get("msg", "")
        except Exception:
            pass
        raise APIError(
            f"Erro HTTP {status} ao acessar a API.\n"
            f"{body}"
        )
    except ValueError:
        raise APIError("A resposta da API não é um JSON válido.")


def buscar_funcionarios(cnpj: str, callback_status=None) -> list[dict]:
    """
    Busca TODOS os funcionários ativos de uma empresa pelo CNPJ.

    Lida com a paginação automática (max 100 por página).
    O `callback_status` é chamado com uma mensagem de progresso (para a UI).

    Retorna lista normalizada:
    [
        {
            "id_funcionario": "123",
            "nome": "João Silva",
            "cpf": "987.654.321-10",
            "cargo": "Eletricista",
            "funcao": "Função",
            "setor": "Obras",
            "data_admissao": "2020-01-15",
            "situacao": "Ativo",
        },
        ...
    ]
    """
    url = get_url_funcionarios()
    todos_funcionarios = []
    pagina = 0
    tamanho = 100  # Máximo permitido pela API

    while True:
        # Formatar CNPJ para o padrão exigido pela API (00.000.000/0000-00)
        c = cnpj
        if len(c) == 14 and "." not in c:
            cnpj_formatado = f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"
        else:
            cnpj_formatado = cnpj

        payload = {
            "paginador": {
                "pagina": str(pagina),
                "tamanho": str(tamanho),
            },
            # Identificação da empresa pelo CNPJ formatado
            "id_empresa": f"CNPJ_CPF_{cnpj_formatado}",
            "situacao": "Ativo",
        }

        if callback_status:
            callback_status(f"🔄  Buscando funcionários (página {pagina + 1})...")

        dados = _fazer_requisicao(url, payload)

        # ── PARSE DA RESPOSTA ──
        # A API retorna: {"resultado": [...], "temProximaPagina": true/false}
        # ou {"resultado": {...}} para um único resultado
        resultado = dados.get("resultado", [])

        # Se for um único objeto, colocar em lista
        if isinstance(resultado, dict):
            resultado = [resultado]

        # Se veio vazio, pode ser erro D001 (retorno em branco) — tratamos como fim
        if not resultado:
            break

        for func in resultado:
            todos_funcionarios.append({
                "id_funcionario": str(func.get("id_funcionario", "")),
                "nome": func.get("nome", ""),
                "cpf": func.get("CPF", ""),
                "cargo": func.get("cargo", ""),
                "funcao": func.get("funcao", ""),
                "setor": func.get("setor", ""),
                "data_admissao": func.get("data_admissao", ""),
                "situacao": func.get("situacao", "Ativo"),
            })

        # Verifica se tem próxima página
        tem_proxima = dados.get("temProximaPagina", False)
        if not tem_proxima:
            break

        pagina += 1

        # Pequena pausa para respeitar rate limit (60 req/min)
        time.sleep(0.5)

    if not todos_funcionarios:
        raise APIError(
            f"Nenhum funcionário ativo encontrado para o CNPJ {cnpj}.\n"
            "Verifique se o CNPJ está correto e se existem funcionários ativos no SGG."
        )

    return todos_funcionarios

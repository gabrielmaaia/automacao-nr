"""
api.py — Requisições à API SGG (endpoint /funcionario/).

Busca todos os funcionários ativos de uma empresa identificada pelo CNPJ ou CPF,
lidando automaticamente com a paginação da API (máximo 100 por página).

Referência da API: https://app.sgg.net.br/api/v3/doc/
"""

import json
import re
import time
from typing import Any

import requests

from config import get_headers, get_url_funcionarios


class APIError(Exception):
    """Exceção customizada para erros de API."""
    pass


def _fazer_requisicao(url: str, payload: dict) -> Any:
    """
    Faz um GET autenticado com filtros e paginação via query params.

    Por causa das regras do CloudFront/WAF, o SGG bloqueia GET com body JSON. A
    API aceita os mesmos parâmetros na URL, então enviamos tudo em `params`.
    """
    try:
        params = payload
        if isinstance(payload, dict) and "paginador" in payload:
            paginador = payload["paginador"]
            if isinstance(paginador, dict):
                params = {
                    "paginador": json.dumps(paginador),
                    "id_empresa": payload.get("id_empresa", ""),
                    "situacao": payload.get("situacao", ""),
                }

        response = requests.get(
            url,
            headers=get_headers(),
            params=params,
            timeout=60,
        )
        response.raise_for_status()
        dados = response.json()

        if isinstance(dados, dict):
            status_code = dados.get("statusCode")
            status_msg = dados.get("statusMsg") or dados.get("msg")
            if status_code and status_code not in ("D000", "D001"):
                raise APIError(
                    f"Erro da API SGG ({status_code}):\n"
                    f"{status_msg or 'Sem mensagem de detalhe'}"
                )

        return dados

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
        if status == 403:
            raise APIError(
                "Acesso bloqueado pelo SGG/CloudFront (403).\n"
                "A chamada foi ajustada para usar query params em vez de body JSON "
                "em GET. Verifique a chave da API e tente novamente."
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


def buscar_funcionarios(documento: str, callback_status=None) -> list[dict]:
    """
    Busca TODOS os funcionários ativos de uma empresa pelo CNPJ ou CPF.

    Lida com a paginação automática (max 100 por página).
    O `callback_status` é chamado com uma mensagem de progresso (para a UI).
    """
    url = get_url_funcionarios()
    todos_funcionarios = []
    pagina = 0
    tamanho = 100

    c = re.sub(r"\D", "", documento)
    if len(c) == 14:
        doc_formatado = f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"
    elif len(c) == 11:
        doc_formatado = f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:11]}"
    else:
        doc_formatado = documento

    while True:
        payload = {
            "paginador": {
                "pagina": str(pagina),
                "tamanho": str(tamanho),
            },
            "id_empresa": f"CNPJ_CPF_{doc_formatado}",
            "situacao": "Ativo",
        }

        if callback_status:
            callback_status(f"🔄  Buscando funcionários (página {pagina + 1})...")

        dados = _fazer_requisicao(url, payload)

        resultado = dados.get("resultado", [])
        if isinstance(resultado, dict):
            resultado = [resultado]

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

        tem_proxima = dados.get("temProximaPagina", False)
        if not tem_proxima:
            break

        pagina += 1
        time.sleep(0.5)

    if not todos_funcionarios:
        raise APIError(
            f"Nenhum funcionário ativo encontrado para o CNPJ/CPF {doc_formatado}.\n"
            "Verifique se o CNPJ/CPF está correto e se existem funcionários ativos no SGG."
        )

    return todos_funcionarios


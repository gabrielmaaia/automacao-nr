"""
config.py — Configuração da API SGG com autenticação Basic Auth.

A API SGG usa Basic Auth com uma peculiaridade:
  - Username = chave de API (32 chars alfanuméricos)
  - Password = vazio (apenas ":")
  - Codificado em Base64 e enviado no header Authorization

Referência: https://app.sgg.net.br/api/v3/doc/
"""

import os
import sys
import base64

from dotenv import load_dotenv

# ============================================================
# Localiza o .env tanto em desenvolvimento quanto no .exe
# ============================================================
if getattr(sys, "frozen", False):
    # Quando o PyInstaller embutir o arquivo, ele ficará na pasta temporária sys._MEIPASS
    _BASE_DIR = sys._MEIPASS
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))

_ENV_PATH = os.path.join(_BASE_DIR, ".env")
load_dotenv(_ENV_PATH)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

# Chave de API obtida pelo sistema SGG
API_KEY: str = os.getenv("SGG_API_KEY", "")

# URL base da API SGG v3
BASE_URL: str = "https://app.sgg.net.br/api/v3"

# Limite de requisições por minuto:
#   60 req/min entre 05h-20h
#  120 req/min entre 20h-05h
# O código deve respeitar a paginação (max 100 por página)


def _build_basic_auth_token() -> str:
    """
    Monta o token Basic Auth no formato da API SGG.

    Formato: Base64("{api_key}:")
    A senha fica em branco, separada da chave apenas por ":"
    """
    credentials = f"{API_KEY}:"
    encoded = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")
    return f"Basic {encoded}"


def get_headers() -> dict:
    """Retorna headers para requisições autenticadas à API SGG."""
    return {
        "Authorization": _build_basic_auth_token(),
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def get_url_funcionarios() -> str:
    """Retorna a URL do endpoint de funcionários."""
    # IMPORTANTE: a API exige a barra '/' no final da URL
    return f"{BASE_URL}/funcionario/"

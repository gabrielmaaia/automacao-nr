"""
processor.py — Cruzamento de dados e geração da Matriz de Treinamentos.

Recebe:
  - Lista de funcionários (da API SGG)
  - Arquivo Excel "Plano de Ação" (treinamentos exigidos gerais)
  - Arquivo Excel "Treinamentos Realizados" (histórico)

Gera:
  - Planilha Excel formatada estilo MATRIZ (1 linha por funcionário, 
    treinamentos nas colunas).
"""

import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class ProcessorError(Exception):
    pass


def _obter_pasta_downloads() -> str:
    pasta = str(Path.home() / "Downloads")
    os.makedirs(pasta, exist_ok=True)
    return pasta


def _normalizar_texto(texto: str) -> str:
    if pd.isna(texto) or texto is None:
        return ""
    return str(texto).strip().upper()


def _carregar_plano_de_acao(caminho_excel: str) -> list[str]:
    try:
        df = pd.read_excel(caminho_excel)
    except Exception as e:
        raise ProcessorError(f"Erro ao ler o arquivo Plano de Ação:\n{e}")

    col_acao = next((col for col in df.columns if "Ação" in col or "Ação" in col or "Acao" in col or "Ao" in col), None)
    if not col_acao:
        raise ProcessorError("Coluna 'Ação' não encontrada no arquivo de Plano de Ação.")

    treinamentos = df[col_acao].dropna().unique().tolist()
    
    # FILTRO: Considerar apenas treinamentos que contêm "NR" no nome
    nrs = [t for t in treinamentos if "NR" in str(t).upper()]
    
    if not nrs:
        raise ProcessorError("Nenhuma ação contendo 'NR' foi encontrada na planilha do Plano de Ação.")
        
    return sorted(nrs)


def _carregar_treinamentos_realizados(caminho_excel: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(caminho_excel)
    except Exception as e:
        raise ProcessorError(f"Erro ao ler o arquivo de Treinamentos Realizados:\n{e}")

    colunas = df.columns
    col_func = next((col for col in colunas if "Funcion" in col), None)
    col_trein = next((col for col in colunas if "Treinamento" in col and "C" not in col), None)
    col_data = next((col for col in colunas if "Data" in col and "Pr" not in col), None)
    col_validade = next((col for col in colunas if "Data Pr" in col), None)

    if not col_func or not col_trein:
        raise ProcessorError("Colunas 'Funcionário' ou 'Treinamento' não encontradas.")

    df = df.rename(columns={
        col_func: "Funcionario",
        col_trein: "Treinamento",
        col_data: "Data Realizacao",
        col_validade: "Data Validade"
    })

    return df


def _calcular_status_e_texto(row: pd.Series) -> str:
    if pd.isna(row.get("Data Realizacao")) or row.get("Data Realizacao") == "":
        return "Pendente"

    validade = row.get("Data Validade")
    texto_base = "Concluído"
    
    # Formata a validade para mostrar na celula
    val_str = ""
    try:
        if isinstance(validade, str):
            val_date = pd.to_datetime(validade, format="%d/%m/%Y", errors="coerce")
        else:
            val_date = pd.to_datetime(validade, errors="coerce")

        if pd.notna(val_date):
            val_str = f" (Vence: {val_date.strftime('%d/%m/%Y')})"
            if val_date >= pd.Timestamp.now().normalize():
                return f"Concluído{val_str}"
            else:
                return f"Vencido{val_str}"
    except Exception:
        pass

    return f"Concluído{val_str}"


def _extrair_codigo_nr(texto: str) -> str:
    match = re.search(r'(?i)(NR[\s-]*\d+)', str(texto))
    if match:
        return match.group(1).upper().replace(" ", "").replace("-", "")
    return ""

def processar_planilha(funcionarios: list[dict], caminho_plano_acao: str, caminho_realizados: str, documento: str) -> str:
    df_func = pd.DataFrame(funcionarios)
    treinamentos_exigidos = _carregar_plano_de_acao(caminho_plano_acao)
    df_realizados = _carregar_treinamentos_realizados(caminho_realizados)

    if df_func.empty:
        raise ProcessorError("Nenhum funcionário retornado pela API.")

    df_func["nome_norm"] = df_func["nome"].apply(_normalizar_texto)
    
    if not df_realizados.empty:
        df_realizados["nome_norm"] = df_realizados["Funcionario"].apply(_normalizar_texto)
        df_realizados["treinamento_norm"] = df_realizados["Treinamento"].apply(_normalizar_texto)
        df_realizados["nr_code"] = df_realizados["Treinamento"].apply(_extrair_codigo_nr)
        df_realizados["Data_Sort"] = pd.to_datetime(df_realizados["Data Realizacao"], format="%d/%m/%Y", errors="coerce")
        df_real_ultimo = (
            df_realizados
            .sort_values("Data_Sort", ascending=False)
            .drop_duplicates(subset=["nome_norm", "nr_code"], keep="first")
        )

    # Cria a estrutura de dados (1 linha por funcionário)
    dados_matriz = []
    
    for _, func in df_func.iterrows():
        linha = {
            "Setor": func.get("setor", ""),
            "Cargo / Função": func.get("cargo", ""),
            "Nome do Funcionário": str(func["nome"]).strip(),
            "CPF": func.get("cpf", ""),
        }
        
        nome_n = func["nome_norm"]
        
        for trein in treinamentos_exigidos:
            nr_code_exigido = _extrair_codigo_nr(trein)
            
            # Buscar histórico deste funcionário + NR
            if not df_realizados.empty and nr_code_exigido:
                match = df_real_ultimo[
                    (df_real_ultimo["nome_norm"] == nome_n) & 
                    (df_real_ultimo["nr_code"] == nr_code_exigido)
                ]
                if not match.empty:
                    row_hist = match.iloc[0]
                    status = _calcular_status_e_texto(row_hist)
                else:
                    status = "Pendente"
            else:
                status = "Pendente"
                
            # Extrair apenas o "NR XX" para o cabeçalho não ficar gigante
            cabecalho_curto = nr_code_exigido if nr_code_exigido else str(trein)[:15]
                
            linha[cabecalho_curto] = status
            
        dados_matriz.append(linha)

    df_final = pd.DataFrame(dados_matriz)
    # Ordenar ignorando case e espaços para garantir ordem alfabética perfeita
    df_final_sorted = df_final.copy()
    df_final_sorted["_sort_key"] = df_final_sorted["Nome do Funcionário"].astype(str).str.strip().str.upper()
    df_final_sorted.sort_values(by=["_sort_key"], inplace=True)
    df_final_sorted.drop(columns=["_sort_key"], inplace=True)
    df_final = df_final_sorted

    # Criar DataFrame de pendências (Formato Lista)
    dados_pendencias = []
    for linha in dados_matriz:
        for key, valor in linha.items():
            if key not in ["Setor", "Cargo / Função", "Nome do Funcionário", "CPF"]:
                if valor.startswith("Pendente") or valor.startswith("Vencido"):
                    dados_pendencias.append({
                        "Setor": linha["Setor"],
                        "Cargo / Função": linha["Cargo / Função"],
                        "Nome do Funcionário": linha["Nome do Funcionário"],
                        "CPF": linha["CPF"],
                        "Treinamento Pendente": key,
                        "Status": valor
                    })
    
    df_pendencias = pd.DataFrame(dados_pendencias)
    if not df_pendencias.empty:
        df_pendencias["_sort_key"] = df_pendencias["Nome do Funcionário"].astype(str).str.strip().str.upper()
        df_pendencias.sort_values(by=["_sort_key", "Treinamento Pendente"], inplace=True)
        df_pendencias.drop(columns=["_sort_key"], inplace=True)

    doc_limpo = re.sub(r"\D", "", str(documento))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"Matriz_Treinamentos_{doc_limpo}_{timestamp}.xlsx"
    caminho_arquivo = os.path.join(_obter_pasta_downloads(), nome_arquivo)

    # Escrever em multiplas abas
    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Matriz NRs")
        df_pendencias.to_excel(writer, index=False, sheet_name="Pendências")

    _formatar_excel(caminho_arquivo)

    return caminho_arquivo


def _formatar_excel(caminho: str) -> None:
    wb = load_workbook(caminho)

    COR_HEADER_BG = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    COR_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    # Estilos para os status
    COR_CONCLUIDO = Font(name="Calibri", size=10, color="1B7A2B")
    FILL_CONCLUIDO = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    
    COR_PENDENTE = Font(name="Calibri", size=10, color="CC6600")
    FILL_PENDENTE = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    COR_VENCIDO = Font(name="Calibri", size=10, bold=True, color="CC0000")
    FILL_VENCIDO = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
    
    BORDA = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"),
    )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Formatar cabeçalho
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = COR_HEADER_BG
            cell.font = COR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDA

        # Formatar dados
        for row_idx in range(2, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = BORDA
                
                valor = str(cell.value) if cell.value else ""
                
                if valor.startswith("Concluído"):
                    cell.font = COR_CONCLUIDO
                    cell.fill = FILL_CONCLUIDO
                elif valor.startswith("Pendente"):
                    cell.font = COR_PENDENTE
                    cell.fill = FILL_PENDENTE
                elif valor.startswith("Vencido"):
                    cell.font = COR_VENCIDO
                    cell.fill = FILL_VENCIDO
                else:
                    cell.font = Font(name="Calibri", size=10)

        # Ajustar largura das colunas
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            header_val = str(ws.cell(row=1, column=col_idx).value)
            
            if header_val == "CPF":
                ws.column_dimensions[col_letter].width = 16
            elif header_val == "Nome do Funcionário":
                ws.column_dimensions[col_letter].width = 30
            elif header_val in ["Setor", "Cargo / Função", "Treinamento Pendente"]:
                ws.column_dimensions[col_letter].width = 25
            elif header_val == "Status":
                ws.column_dimensions[col_letter].width = 20
            else:
                # Colunas de treinamentos (NRs) na aba Matriz
                ws.column_dimensions[col_letter].width = 14

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    wb.save(caminho)

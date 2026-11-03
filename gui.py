"""
gui.py — Interface gráfica com CustomTkinter (tema escuro/moderno).

Tela com:
  1. Campo de CNPJ (com máscara)
  2. Seletor do arquivo Excel "Plano de Ação"
  3. Seletor do arquivo Excel "Treinamentos Realizados"
  4. Botão para gerar a Matriz de Treinamentos
  5. Feedback visual com barra de progresso
"""

import os
import re
import subprocess
import sys
import threading
from tkinter import filedialog

import customtkinter as ctk

from api import buscar_funcionarios, APIError
from processor import processar_planilha, ProcessorError


# ============================================================
# CONFIGURAÇÃO DO TEMA
# ============================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COR_FUNDO = "#0D1117"
COR_CARD = "#161B22"
COR_BORDA = "#30363D"
COR_PRIMARIA = "#2F81F7"
COR_PRIMARIA_HOVER = "#1A6FE0"
COR_SUCESSO = "#2EA043"
COR_ERRO = "#F85149"
COR_TEXTO = "#E6EDF3"
COR_TEXTO_SECUNDARIO = "#8B949E"
COR_INPUT_BG = "#0D1117"


class MatrizApp(ctk.CTk):
    """Janela principal do aplicativo."""

    def __init__(self):
        super().__init__()

        self.title("Sistema Organização Nr — Igor")
        self.geometry("640x780")
        self.minsize(580, 720)
        self.configure(fg_color=COR_FUNDO)
        self.resizable(True, True)

        self._caminho_arquivo_saida: str = ""
        self._caminho_plano_acao: str = ""
        self._caminho_realizados: str = ""
        self._processando: bool = False

        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (640 // 2)
        y = (self.winfo_screenheight() // 2) - (780 // 2)
        self.geometry(f"+{x}+{y}")

        self._criar_widgets()

    def _criar_widgets(self):
        self.card = ctk.CTkFrame(
            self,
            fg_color=COR_CARD,
            corner_radius=16,
            border_width=1,
            border_color=COR_BORDA,
        )
        self.card.pack(padx=28, pady=(28, 12), fill="both", expand=True)

        # ── Cabeçalho (Ícone + Textos centralizados) ──
        self.lbl_icon = ctk.CTkLabel(
            self.card, text="🛡️", font=ctk.CTkFont(size=44), text_color=COR_TEXTO
        )
        # padx=(28, 0) empurra o emoji levemente para a direita para compensar o desvio visual da fonte
        self.lbl_icon.pack(pady=(24, 2), padx=(42, 0), anchor="center")

        self.lbl_titulo = ctk.CTkLabel(
            self.card, text="Matriz de Treinamentos",
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
            text_color=COR_TEXTO,
        )
        self.lbl_titulo.pack(pady=(0, 2), anchor="center")

        self.lbl_subtitulo = ctk.CTkLabel(
            self.card, text="Segurança do Trabalho — Sistema SGG",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COR_TEXTO_SECUNDARIO,
        )
        self.lbl_subtitulo.pack(pady=(0, 16), anchor="center")

        ctk.CTkFrame(self.card, height=1, fg_color=COR_BORDA).pack(fill="x", padx=28, pady=(0, 16))

        # ── SEÇÃO 1: CNPJ / CPF ──
        self._label_secao("① CNPJ ou CPF da Empresa")
        self.entry_cnpj = ctk.CTkEntry(
            self.card, placeholder_text="CNPJ (14 dig) ou CPF (11 dig)",
            font=ctk.CTkFont(family="Consolas", size=15),
            height=44, corner_radius=10, border_width=1,
            border_color=COR_BORDA, fg_color=COR_INPUT_BG, text_color=COR_TEXTO,
        )
        self.entry_cnpj.pack(padx=36, fill="x", pady=(4, 2))
        self.entry_cnpj.bind("<KeyRelease>", self._mascara_documento)
        self._label_dica("Digite 14 dígitos (CNPJ) ou 11 dígitos (CPF)")

        # ── SEÇÃO 2: PLANO DE AÇÃO ──
        self._label_secao("② Planilha: Plano de Ação")
        frame_acao = ctk.CTkFrame(self.card, fg_color="transparent")
        frame_acao.pack(padx=36, fill="x", pady=(4, 2))

        self.btn_acao = ctk.CTkButton(
            frame_acao, text="📄 Selecionar", font=ctk.CTkFont(family="Segoe UI", size=13),
            height=36, width=120, corner_radius=8, fg_color=COR_BORDA, hover_color="#484F58",
            text_color=COR_TEXTO, command=self._selecionar_plano_acao,
        )
        self.btn_acao.pack(side="left")

        self.lbl_acao = ctk.CTkLabel(
            frame_acao, text="  Nenhum arquivo", font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COR_TEXTO_SECUNDARIO, anchor="w",
        )
        self.lbl_acao.pack(side="left", padx=(10, 0), fill="x", expand=True)

        # ── SEÇÃO 3: TREINAMENTOS REALIZADOS ──
        self._label_secao("③ Planilha: Treinamentos Realizados")
        frame_realizados = ctk.CTkFrame(self.card, fg_color="transparent")
        frame_realizados.pack(padx=36, fill="x", pady=(4, 2))

        self.btn_realizados = ctk.CTkButton(
            frame_realizados, text="📄 Selecionar", font=ctk.CTkFont(family="Segoe UI", size=13),
            height=36, width=120, corner_radius=8, fg_color=COR_BORDA, hover_color="#484F58",
            text_color=COR_TEXTO, command=self._selecionar_realizados,
        )
        self.btn_realizados.pack(side="left")

        self.lbl_realizados = ctk.CTkLabel(
            frame_realizados, text="  Nenhum arquivo", font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COR_TEXTO_SECUNDARIO, anchor="w",
        )
        self.lbl_realizados.pack(side="left", padx=(10, 0), fill="x", expand=True)

        ctk.CTkFrame(self.card, height=1, fg_color=COR_BORDA).pack(fill="x", padx=28, pady=(20, 16))

        # ── BOTÃO PRINCIPAL ──
        self.btn_gerar = ctk.CTkButton(
            self.card, text="⚡  Gerar Matriz de Treinamentos",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            height=50, corner_radius=10, fg_color=COR_PRIMARIA, hover_color=COR_PRIMARIA_HOVER,
            command=self._ao_clicar_gerar,
        )
        self.btn_gerar.pack(padx=36, fill="x", pady=(0, 12))

        # Status + Progresso
        self.lbl_status = ctk.CTkLabel(
            self.card, text="", font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COR_TEXTO_SECUNDARIO, wraplength=500, justify="center",
        )
        self.lbl_status.pack(padx=36, pady=(0, 4))

        self.progress = ctk.CTkProgressBar(
            self.card, height=5, corner_radius=3, fg_color=COR_BORDA,
            progress_color=COR_PRIMARIA, mode="indeterminate",
        )

        # Botões resultado
        self.btn_abrir = ctk.CTkButton(
            self.card, text="📂  Abrir Arquivo", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            height=42, corner_radius=10, fg_color=COR_SUCESSO, hover_color="#26833A", command=self._abrir_arquivo,
        )
        self.btn_abrir_pasta = ctk.CTkButton(
            self.card, text="📁  Abrir Pasta Downloads", font=ctk.CTkFont(family="Segoe UI", size=12),
            height=36, corner_radius=10, fg_color="transparent", hover_color=COR_BORDA, border_width=1,
            border_color=COR_BORDA, text_color=COR_TEXTO_SECUNDARIO, command=self._abrir_pasta,
        )

        # Rodapé
        self.lbl_rodape = ctk.CTkLabel(
            self, text="v3.0 • API SGG + Planilhas Exportadas",
            font=ctk.CTkFont(family="Segoe UI", size=10), text_color=COR_TEXTO_SECUNDARIO,
        )
        self.lbl_rodape.pack(side="bottom", pady=(0, 8))

    def _label_secao(self, texto: str):
        ctk.CTkLabel(
            self.card, text=texto, font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=COR_TEXTO, anchor="w",
        ).pack(padx=38, anchor="w", pady=(8, 0))

    def _label_dica(self, texto: str):
        ctk.CTkLabel(
            self.card, text=texto, font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COR_TEXTO_SECUNDARIO, anchor="w",
        ).pack(padx=40, anchor="w", pady=(0, 6))

    def _mascara_documento(self, event=None):
        texto = self.entry_cnpj.get()
        apenas_digitos = re.sub(r"\D", "", texto)[:14]
        
        formatado = ""
        if len(apenas_digitos) <= 11:
            # Formato CPF (000.000.000-00)
            for i, d in enumerate(apenas_digitos):
                if i in (3, 6):
                    formatado += "."
                elif i == 9:
                    formatado += "-"
                formatado += d
        else:
            # Formato CNPJ (00.000.000/0000-00)
            for i, d in enumerate(apenas_digitos):
                if i in (2, 5):
                    formatado += "."
                elif i == 8:
                    formatado += "/"
                elif i == 12:
                    formatado += "-"
                formatado += d

        self.entry_cnpj.delete(0, "end")
        self.entry_cnpj.insert(0, formatado)

    def _selecionar_plano_acao(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Plano de Ação",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self._caminho_plano_acao = caminho
            self.lbl_acao.configure(text=f"  ✅ {os.path.basename(caminho)}", text_color=COR_SUCESSO)

    def _selecionar_realizados(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Treinamentos Realizados",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self._caminho_realizados = caminho
            self.lbl_realizados.configure(text=f"  ✅ {os.path.basename(caminho)}", text_color=COR_SUCESSO)

    def _extrair_documento(self) -> str | None:
        digitos = re.sub(r"\D", "", self.entry_cnpj.get())
        return digitos if len(digitos) in (11, 14) else None

    def _ao_clicar_gerar(self):
        if self._processando: return

        documento = self._extrair_documento()
        if not documento:
            self._exibir_status("⚠️  CNPJ ou CPF inválido. Digite 11 (CPF) ou 14 dígitos (CNPJ).", cor=COR_ERRO)
            return
        if not self._caminho_plano_acao:
            self._exibir_status("⚠️  Selecione a planilha de Plano de Ação.", cor=COR_ERRO)
            return
        if not self._caminho_realizados:
            self._exibir_status("⚠️  Selecione a planilha de Treinamentos Realizados.", cor=COR_ERRO)
            return

        self.btn_abrir.pack_forget()
        self.btn_abrir_pasta.pack_forget()
        self._processando = True
        self.btn_gerar.configure(state="disabled", text="Processando...")
        self._exibir_status("🔄  Conectando à API do SGG...", cor=COR_TEXTO_SECUNDARIO)
        self.progress.pack(padx=36, fill="x", pady=(0, 8))
        self.progress.start()

        threading.Thread(target=self._executar_pipeline, args=(documento,), daemon=True).start()

    def _executar_pipeline(self, documento: str):
        try:
            def status_cb(msg): self.after(0, self._exibir_status, msg, COR_TEXTO_SECUNDARIO)
            
            funcionarios = buscar_funcionarios(documento, callback_status=status_cb)
            
            self.after(0, self._exibir_status, f"🔄  Cruzando {len(funcionarios)} funcionários com as planilhas...", COR_TEXTO_SECUNDARIO)
            
            caminho = processar_planilha(
                funcionarios, self._caminho_plano_acao, self._caminho_realizados, documento
            )

            self._caminho_arquivo_saida = caminho
            self.after(0, self._finalizar_sucesso, caminho, len(funcionarios))

        except (APIError, ProcessorError) as e:
            self.after(0, self._finalizar_erro, str(e))
        except Exception as e:
            self.after(0, self._finalizar_erro, f"Erro inesperado:\n{e}")

    def _exibir_status(self, texto: str, cor: str):
        self.lbl_status.configure(text=texto, text_color=cor)

    def _finalizar_sucesso(self, caminho: str, total_func: int):
        self.progress.stop()
        self.progress.pack_forget()
        self._processando = False
        self.btn_gerar.configure(state="normal", text="⚡  Gerar Matriz de Treinamentos")
        self._exibir_status(f"✅  Planilha gerada! ({total_func} funcionários)\n📄  {os.path.basename(caminho)}", cor=COR_SUCESSO)
        self.btn_abrir.pack(padx=36, fill="x", pady=(8, 6))
        self.btn_abrir_pasta.pack(padx=36, fill="x", pady=(0, 12))

    def _finalizar_erro(self, mensagem: str):
        self.progress.stop()
        self.progress.pack_forget()
        self._processando = False
        self.btn_gerar.configure(state="normal", text="⚡  Gerar Matriz de Treinamentos")
        self._exibir_status(f"❌  {mensagem}", cor=COR_ERRO)

    def _abrir_arquivo(self):
        if self._caminho_arquivo_saida and os.path.exists(self._caminho_arquivo_saida):
            os.startfile(self._caminho_arquivo_saida) if sys.platform == "win32" else subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", self._caminho_arquivo_saida])

    def _abrir_pasta(self):
        pasta = os.path.dirname(self._caminho_arquivo_saida) if self._caminho_arquivo_saida else str(__import__("pathlib").Path.home() / "Downloads")
        os.startfile(pasta) if sys.platform == "win32" else subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", pasta])


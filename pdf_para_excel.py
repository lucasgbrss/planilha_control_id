import os
import email as email_lib
import pdfplumber
from bs4 import BeautifulSoup
from pathlib import Path
from openpyxl import Workbook
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from datetime import datetime
import threading
import json
import os


class PdfToExcelApp:
    """Aplicativo moderno para converter folhas de ponto em Excel."""

    def __init__(self, root):
        self.root = root
        self.root.title("Control ID Reader - Conversor de Ponto")
        self.root.geometry("820x680")
        self.root.resizable(True, True)
        self.root.minsize(700, 580)

        # Tema padrão
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Aplicar ícone da janela
        self._aplicar_icone()

        # Centralizar janela na tela
        self.centralizar_janela()

        # Variáveis de estado
        self.pdf_paths = []
        self.processamento_ativo = False
        self._tema_atual = "dark"

        # Diretórios padrão (serão sobrescritos pelas configurações salvas)
        self.dir_abrir_pdf = str(Path.home() / "Downloads")
        self.dir_salvar_excel = str(Path.home() / "Documents")

        # Criar interface
        self.criar_interface()

        # Carregar configurações salvas
        self.carregar_configuracoes()

    def _aplicar_icone(self):
        """Aplica o ícone da janela a partir do arquivo .ico na pasta do script."""
        try:
            ico_path = Path(__file__).parent / "control_id_reader.ico"
            if ico_path.exists():
                self.root.iconbitmap(str(ico_path))
        except Exception:
            pass  # Silencioso — ícone é cosmético, não deve travar o app

    def centralizar_janela(self):
        """Centraliza a janela na tela."""
        self.root.update_idletasks()
        largura = self.root.winfo_width()
        altura = self.root.winfo_height()
        tela_largura = self.root.winfo_screenwidth()
        tela_altura = self.root.winfo_screenheight()
        x = (tela_largura // 2) - (largura // 2)
        y = (tela_altura // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{x}+{y}")

    def alternar_tema(self):
        """Alterna entre dark e light mode."""
        if self._tema_atual == "dark":
            ctk.set_appearance_mode("light")
            self._tema_atual = "light"
            self.btn_tema.configure(text="☀️  Light")
        else:
            ctk.set_appearance_mode("dark")
            self._tema_atual = "dark"
            self.btn_tema.configure(text="🌙  Dark")
        self.root.after(50, self._atualizar_cores_log)

    def _atualizar_cores_log(self):
        """Sincroniza as cores do tk.Text e tk.Listbox com o tema CTk atual."""
        is_dark = self._tema_atual == "dark"
        bg      = "#1e1e1e" if is_dark else "#f0f0f0"
        fg      = "#d4d4d4" if is_dark else "#222222"
        sel_bg  = "#264f78" if is_dark else "#1f6aa5"

        self.log_text.configure(bg=bg, fg=fg,
                                insertbackground=fg, selectbackground=sel_bg)
        self.log_text.tag_config("info",       foreground="#d4d4d4" if is_dark else "#333333")
        self.log_text.tag_config("success",    foreground="#4ec94e" if is_dark else "#1e7e34")
        self.log_text.tag_config("error",      foreground="#f47070" if is_dark else "#c0392b")
        self.log_text.tag_config("processing", foreground="#f0c060" if is_dark else "#d68910")

        self.lista_arquivos.configure(
            bg=bg, fg=fg,
            selectbackground=sel_bg, selectforeground="white"
        )

    def criar_interface(self):
        """Cria todos os componentes da interface."""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Container principal
        main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)  # log expande verticalmente

        # ── Cabeçalho ───────────────────────────────────────────
        header = ctk.CTkFrame(main_frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)

        ctk.CTkLabel(header, text="Control ID Reader",
                     font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
                     ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(header, text="Conversor de Folhas de Ponto para Excel",
                     font=ctk.CTkFont(family="Segoe UI", size=12),
                     text_color=("gray40", "gray70")
                     ).grid(row=1, column=0, sticky="w")

        # Botão tema no canto superior direito
        self.btn_tema = ctk.CTkButton(
            header, text="🌙  Dark", width=100, height=30,
            fg_color=("gray80", "gray25"), text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray35"),
            command=self.alternar_tema
        )
        self.btn_tema.grid(row=0, column=1, rowspan=2, sticky="e")

        # Seções
        self.criar_secao_selecao(main_frame, 1)
        self.criar_secao_status(main_frame, 2)
        self.criar_secao_acoes(main_frame, 3)
    def criar_secao_selecao(self, parent, row):
        """Cria a seção para seleção de arquivos."""
        frame = ctk.CTkFrame(parent)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Seleção de Arquivos",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 6))

        # Lista de arquivos (tk.Listbox ainda não tem substituto CTk nativo)
        list_frame = ctk.CTkFrame(frame, fg_color="transparent")
        list_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 4))
        list_frame.columnconfigure(0, weight=1)

        self.lista_arquivos = tk.Listbox(
            list_frame, height=5, selectmode=tk.EXTENDED,
            font=("Segoe UI", 9), relief="flat", borderwidth=0,
            activestyle="none", selectbackground="#1f6aa5",
            selectforeground="white"
        )
        self.lista_arquivos.grid(row=0, column=0, sticky="ew")

        scrollbar = ctk.CTkScrollbar(list_frame, command=self.lista_arquivos.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.lista_arquivos.configure(yscrollcommand=scrollbar.set)

        # Botões
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="w", padx=12, pady=(4, 4))

        ctk.CTkButton(btn_frame, text="+ Adicionar Arquivos", width=160,
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=self.adicionar_pdfs
                      ).grid(row=0, column=0, padx=(0, 6))

        ctk.CTkButton(btn_frame, text="✕ Remover Selecionados", width=170,
                      fg_color="#e67e22", hover_color="#ca6f1e",
                      command=self.remover_selecionados
                      ).grid(row=0, column=1, padx=(0, 6))

        ctk.CTkButton(btn_frame, text="🗑  Limpar Tudo", width=130,
                      fg_color="#c0392b", hover_color="#a93226",
                      command=self.limpar_lista
                      ).grid(row=0, column=2)

        self.info_arquivos = ctk.CTkLabel(
            frame, text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=11), text_color=("gray40", "gray60")
        )
        self.info_arquivos.grid(row=3, column=0, sticky="w", padx=14, pady=(0, 10))

    def criar_secao_status(self, parent, row):
        """Cria a seção de log/status."""
        frame = ctk.CTkFrame(parent)
        frame.grid(row=row, column=0, sticky="nsew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ctk.CTkLabel(frame, text="Log de Processamento",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 6))

        self.log_text = tk.Text(
            frame, height=9, wrap=tk.WORD,
            font=("Consolas", 9), relief="flat", borderwidth=0,
            padx=8, pady=6
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        log_scroll = ctk.CTkScrollbar(frame, command=self.log_text.yview)
        log_scroll.grid(row=1, column=1, sticky="ns", pady=(0, 12), padx=(0, 8))
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self._atualizar_cores_log()

    def criar_secao_acoes(self, parent, row):
        """Cria a seção de botões de ação."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", pady=(4, 0))
        frame.columnconfigure(0, weight=1)

        # Barra de progresso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(frame, variable=self.progress_var,
                                               height=10)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, columnspan=3, sticky="ew",
                               pady=(0, 10))

        # Botão principal
        self.btn_processar = ctk.CTkButton(
            frame, text="⚡  Gerar Planilha Excel",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40, fg_color="#2980b9", hover_color="#2471a3",
            command=self.iniciar_processamento
        )
        self.btn_processar.grid(row=1, column=0, sticky="w")

        # Botão secundário
        ctk.CTkButton(
            frame, text="💾  Salvar Configurações",
            height=40, width=180,
            fg_color=("gray75", "gray30"), text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray40"),
            command=self.salvar_configuracoes
        ).grid(row=1, column=1, padx=(10, 0), sticky="w")

        self.status_final = ctk.CTkLabel(
            frame, text="",
            font=ctk.CTkFont(size=11), text_color=("gray40", "gray60")
        )
        self.status_final.grid(row=2, column=0, columnspan=3,
                               sticky="w", pady=(8, 0))

    def adicionar_pdfs(self):
        """Abre dialog para selecionar arquivos PDF ou MHTML."""
        arquivo_types = [
            ("Arquivos suportados", "*.pdf *.mhtml"),
            ("Arquivos PDF", "*.pdf"),
            ("Arquivos MHTML", "*.mhtml"),
        ]
        files = filedialog.askopenfilenames(
            title="Selecionar Arquivos PDF ou MHTML",
            filetypes=arquivo_types,
            initialdir=self.dir_abrir_pdf
        )

        if files:
            for f in files:
                if f not in self.pdf_paths:
                    self.pdf_paths.append(f)
                    self.lista_arquivos.insert(tk.END, Path(f).name)
            self.info_arquivos.configure(text=f"{len(self.pdf_paths)} arquivo(s) selecionado(s)")

            # Atualizar o último diretório usado
            self.dir_abrir_pdf = str(Path(files[0]).parent)
            self.salvar_configuracoes(silencioso=True)

    def remover_selecionados(self):
        """Remove arquivos selecionados da lista."""
        indices = self.lista_arquivos.curselection()
        if not indices:
            return

        for i in reversed(indices):
            self.lista_arquivos.delete(i)
            self.pdf_paths.pop(i)

        self.info_arquivos.configure(text=f"{len(self.pdf_paths)} arquivo(s) selecionado(s)")

    def limpar_lista(self):
        """Limpa todos os arquivos da lista."""
        self.pdf_paths.clear()
        self.lista_arquivos.delete(0, tk.END)
        self.info_arquivos.configure(text="Nenhum arquivo selecionado")

    def log(self, mensagem, tipo='info'):
        """Adiciona mensagem ao log de forma thread-safe via root.after."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        def _inserir():
            self.log_text.insert(tk.END, f"[{timestamp}] {mensagem}\n", tipo)
            self.log_text.see(tk.END)
        self.root.after(0, _inserir)

    def _finalizar_processamento(self, todos_dias, total_pdfs, total_dias):
        """Chamado na thread principal ao fim do processamento para abrir o filedialog."""
        self.processamento_ativo = False
        self.btn_processar.configure(state="normal")
        if todos_dias:
            self.salvar_excel(todos_dias)
            self.log(f"\n✅ Processamento concluído!", "success")
            self.log(f"Total: {total_pdfs} arquivo(s), {total_dias} dia(s) registrado(s)", "success")
        else:
            self.log("\n⚠️ Nenhum dado foi extraído dos arquivos.", "error")

    def extrair_info_funcionario(self, texto):
        """Extrai informações do funcionário a partir do texto de um arquivo.
        Retorna (dict, motivo_falha). motivo_falha é None em caso de sucesso."""
        info = {}

        if not texto or not texto.strip():
            return info, "texto do arquivo está vazio"

        # EMPRESA e CNPJ na mesma linha
        match = re.search(r'EMPRESA:\s*(.+?)\s*CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
                         texto, re.IGNORECASE)
        if match:
            info["empresa"] = match.group(1).strip()
            info["cnpj"] = match.group(2).strip()
        else:
            match = re.search(r'EMPRESA:\s*([^\n]+)', texto, re.IGNORECASE)
            if match:
                info["empresa"] = match.group(1).strip()
            match = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto, re.IGNORECASE)
            if match:
                info["cnpj"] = match.group(1).strip()

        # ENDEREÇO
        match = re.search(r'ENDERE[Ç]O:\s*(.+?)(?:\n|$)', texto, re.IGNORECASE)
        if match:
            info["endereco"] = match.group(1).strip()

        # NOME
        match = re.search(r'NOME:\s*(.+?)\s+PIS/PASEP:', texto, re.IGNORECASE)
        if match:
            info["nome"] = match.group(1).strip()

        # PIS/PASEP
        match = re.search(r'PIS/PASEP:\s*(\d+)', texto, re.IGNORECASE)
        if match:
            info["pis"] = match.group(1).strip()

        # ADMISSÃO
        match = re.search(r'ADMISS[Ã]O:\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
        if match:
            info["admissao"] = match.group(1).strip()

        # CPF
        match = re.search(r'CPF:\s*(\d+)', texto, re.IGNORECASE)
        if match:
            info["cpf"] = match.group(1).strip()

        # MATRÍCULA
        match = re.search(r'MATR[Í]CULA:\s*(\d+)', texto, re.IGNORECASE)
        if match:
            info["matricula"] = match.group(1).strip()

        # CENTRO DE CUSTO
        match = re.search(r'CENTRO DE CUSTO:\s*(\S+)', texto, re.IGNORECASE)
        if match:
            info["centro_custo"] = match.group(1).strip()

        # DEPARTAMENTO
        match = re.search(r'DEPARTAMENTO:\s*(\S+)', texto, re.IGNORECASE)
        if match:
            info["departamento"] = match.group(1).strip()

        # CARGO
        match = re.search(r'CARGO:\s*(.+?)(?:\n|$)', texto, re.IGNORECASE)
        if match:
            info["cargo"] = match.group(1).strip()

        # Fix 6: retornar motivo específico se nome não foi encontrado
        if not info.get("nome"):
            tem_empresa = "empresa" in info
            tem_cpf     = "cpf" in info
            if not tem_empresa and not tem_cpf:
                motivo = "nenhum campo reconhecido — formato do arquivo pode ser diferente do esperado"
            elif tem_empresa and not tem_cpf:
                motivo = "empresa identificada mas campo NOME não encontrado no padrão 'NOME: ... PIS/PASEP:'"
            else:
                motivo = "campo NOME não encontrado no padrão esperado"
            return info, motivo

        return info, None

    def extrair_dados_ponto(self, tabelas):
        """Extrai dados de ponto de uma lista de tabelas."""
        dias = []
        dia_atual = None

        for row in tabelas:
            # Detectar nova linha de dia
            if row and len(row) > 0 and row[0]:
                dia_str = str(row[0]).upper()
                if any(mes in dia_str for mes in ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]):
                    dia_atual = str(row[0]).strip()
                    dia = {
                        "dia": dia_atual,
                        "marcacoes": "",
                        "ent1": "", "sai1": "", "ent2": "", "sai2": "",
                        "ent3": "", "sai3": "", "duracao": "", "ch": ""
                    }
                    dias.append(dia)

            if dia_atual and len(dias) > 0:
                idx = len(dias) - 1
                # Marcações
                if len(row) > 1 and row[1]:
                    dias[idx]["marcacoes"] = str(row[1]).strip()
                # Jornada
                if len(row) > 2 and row[2]:
                    dias[idx]["ent1"] = str(row[2]).strip()
                if len(row) > 3 and row[3]:
                    dias[idx]["sai1"] = str(row[3]).strip()
                if len(row) > 4 and row[4]:
                    dias[idx]["ent2"] = str(row[4]).strip()
                if len(row) > 5 and row[5]:
                    dias[idx]["sai2"] = str(row[5]).strip()
                if len(row) > 6 and row[6]:
                    dias[idx]["ent3"] = str(row[6]).strip()
                if len(row) > 7 and row[7]:
                    dias[idx]["sai3"] = str(row[7]).strip()
                # Duração e CH
                if len(row) > 8 and row[8]:
                    dias[idx]["duracao"] = str(row[8]).strip()
                if len(row) > 9 and row[9]:
                    dias[idx]["ch"] = str(row[9]).strip()

        return dias

    # ──────────────────────────────────────────────
    # Extração MHTML
    # ──────────────────────────────────────────────

    def extrair_html_de_mhtml(self, caminho):
        """Abre um arquivo .mhtml e retorna o conteúdo HTML interno como string."""
        with open(caminho, "rb") as f:
            msg = email_lib.message_from_bytes(f.read())
        for part in msg.walk():
            if "html" in part.get_content_type():
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                return payload.decode(charset, errors="replace")
        return None

    def extrair_info_funcionario_mhtml(self, soup):
        """Extrai dados do funcionário a partir do HTML parseado.
        Retorna (dict, motivo_falha). Células têm formato 'CAMPO:Valor'."""
        info = {}

        campo_map_lower = {
            "empresa"        : "empresa",
            "cnpj"           : "cnpj",
            "cei"            : "cei",
            "endereço"       : "endereco",
            "nome"           : "nome",
            "pis/pasep"      : "pis",
            "admissão"       : "admissao",
            "centro de custo": "centro_custo",
            "cpf"            : "cpf",
            "matrícula"      : "matricula",
            "departamento"   : "departamento",
            "cargo"          : "cargo",
        }

        for td in soup.find_all(["td", "th"]):
            texto = td.get_text(separator=" ", strip=True)
            if ":" not in texto:
                continue
            chave, _, valor = texto.partition(":")
            chave_norm = chave.strip().lower()
            valor = valor.strip()
            if chave_norm in campo_map_lower and valor:
                info[campo_map_lower[chave_norm]] = valor

        if not info.get("nome"):
            tem_empresa = "empresa" in info
            if not tem_empresa:
                motivo = "nenhum campo reconhecido — estrutura do MHTML diferente do esperado"
            else:
                motivo = "empresa identificada mas campo NOME não encontrado"
            return info, motivo

        return info, None

    def extrair_dados_ponto_mhtml(self, soup):
        """Extrai registros de ponto da tabela HTML.
        Cada linha-mestre tem 17 células: [0]=data, [3]=ENT1, [4]=SAÍ1,
        [5]=ENT2, [6]=SAÍ2. Linhas de detalhe (3 células) são ignoradas."""
        dias = []
        tables = soup.find_all("table")

        tabela_ponto = None
        for table in tables:
            primeira = table.find("tr")
            if primeira and "DIA" in primeira.get_text():
                tabela_ponto = table
                break

        if not tabela_ponto:
            return dias

        data_re = re.compile(r"(\d{2}/\d{2}/\d{2,4})")

        for tr in tabela_ponto.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]

            if len(cells) < 17:
                continue
            m = data_re.match(cells[0])
            if not m:
                continue

            data_raw  = cells[0]
            data_part = m.group(1)

            try:
                from datetime import datetime as dt
                fmt = "%d/%m/%y" if len(data_part.split("/")[2]) == 2 else "%d/%m/%Y"
                data_fmt = dt.strptime(data_part, fmt).strftime("%d/%m/%Y")
            except ValueError:
                data_fmt = data_part

            dia_semana = data_raw.split("-")[-1].strip() if "-" in data_raw else ""

            ent1    = cells[3].strip()
            sai1    = cells[4].strip()
            ent2    = cells[5].strip()
            sai2    = cells[6].strip()
            duracao = cells[10].strip() if len(cells) > 10 else ""
            ch      = cells[12].strip() if len(cells) > 12 else ""

            # "marcacoes" = resumo visual das batidas (ex: "07:23 14:27")
            marcacoes = f"{ent1} {sai1}".strip() if ent1 or sai1 else ""

            dias.append({
                "dia"     : f"{data_fmt} - {dia_semana}",
                "marcacoes": marcacoes,
                "ent1"    : ent1,
                "sai1"    : sai1,
                "ent2"    : ent2,
                "sai2"    : sai2,
                "ent3"    : "",
                "sai3"    : "",
                "duracao" : duracao,
                "ch"      : ch,
            })

        return dias

    # ──────────────────────────────────────────────

    def processar_pdfs_thread(self):
        """Thread principal de processamento."""
        # Fix 7: flag ativado antes do try para garantir consistência com o botão desabilitado
        self.processamento_ativo = True
        total_pdfs = len(self.pdf_paths)
        total_dias = 0
        todos_dias = []

        try:
            self.log(f"Iniciando processamento de {total_pdfs} arquivo(s)...", "processing")

            for idx, pdf_path in enumerate(self.pdf_paths):
                nome_arquivo = Path(pdf_path).name
                self.log(f"Processando: {nome_arquivo}", "processing")

                try:
                    # Fix 3: verificar se o arquivo ainda existe antes de abrir
                    if not Path(pdf_path).exists():
                        self.log(f"  ✗ Arquivo não encontrado: {nome_arquivo}", "error")
                        continue

                    ext = Path(pdf_path).suffix.lower()

                    if ext == ".mhtml":
                        # ── Fluxo MHTML ──────────────────────────────
                        html_str = self.extrair_html_de_mhtml(pdf_path)
                        if not html_str:
                            self.log(f"  ✗ Não foi possível extrair HTML de '{nome_arquivo}'.", "error")
                            continue

                        soup = BeautifulSoup(html_str, "html.parser")
                        info_funcionario, motivo_falha = self.extrair_info_funcionario_mhtml(soup)

                        if not info_funcionario.get("nome"):
                            self.log(f"  ⚠️ Nenhum funcionário identificado em '{nome_arquivo}'. "
                                     f"Motivo: {motivo_falha}", "error")
                            continue

                        dias_do_pdf = self.extrair_dados_ponto_mhtml(soup)

                    else:
                        # ── Fluxo PDF ────────────────────────────────
                        texto_do_pdf  = ""
                        tabelas_do_pdf = []

                        try:
                            pdf_aberto = pdfplumber.open(pdf_path)
                        except Exception:
                            self.log(f"  ✗ Não foi possível abrir '{nome_arquivo}'. "
                                     f"O arquivo pode estar protegido por senha ou corrompido.", "error")
                            continue

                        with pdf_aberto as pdf:
                            for pagina in pdf.pages:
                                texto_do_pdf += pagina.extract_text() or ""
                                tables = pagina.extract_tables()
                                for table in tables:
                                    if table:
                                        tabelas_do_pdf.extend(table)

                        if not texto_do_pdf.strip():
                            self.log(f"  ✗ '{nome_arquivo}' não contém texto extraível. "
                                     f"Pode estar protegido por senha ou ser um PDF digitalizado (imagem).", "error")
                            continue

                        info_funcionario, motivo_falha = self.extrair_info_funcionario(texto_do_pdf)

                        if not info_funcionario.get("nome"):
                            self.log(f"  ⚠️ Nenhum funcionário identificado em '{nome_arquivo}'. "
                                     f"Motivo: {motivo_falha}", "error")
                            continue

                        dias_do_pdf = self.extrair_dados_ponto(tabelas_do_pdf)

                    for dia in dias_do_pdf:
                        dia["nome"]    = info_funcionario.get("nome", "")
                        dia["empresa"] = info_funcionario.get("empresa", "")
                        dia["cpf"]     = info_funcionario.get("cpf", "")

                    todos_dias.extend(dias_do_pdf)
                    total_dias += len(dias_do_pdf)
                    self.log(f"  ✓ {len(dias_do_pdf)} dia(s) extraído(s)", "success")

                except PermissionError:
                    self.log(f"  ✗ Sem permissão para ler '{nome_arquivo}'. "
                             f"Verifique se o arquivo está aberto em outro programa.", "error")
                except MemoryError:
                    self.log(f"  ✗ Memória insuficiente ao processar '{nome_arquivo}'. "
                             f"Tente processar menos arquivos por vez.", "error")
                except Exception as e:
                    self.log(f"  ✗ Erro inesperado em '{nome_arquivo}': {str(e)}", "error")

                self.root.after(0, lambda: self.progress_bar.set(((idx + 1) / total_pdfs)))

        except Exception as e:
            self.log(f"Erro crítico no processamento: {str(e)}", "error")
        finally:
            # Fix 1: devolver controle à thread principal para abrir o filedialog
            self.root.after(0, lambda: self._finalizar_processamento(todos_dias, total_pdfs, total_dias))

    def salvar_excel(self, dados):
        """Salva os dados em um arquivo Excel."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_sugerido = f"Ponto_{len(self.pdf_paths)}PDFs_{timestamp}.xlsx"

        excel_path = filedialog.asksaveasfilename(
            title="Salvar planilha Excel",
            defaultextension=".xlsx",
            initialfile=nome_sugerido,
            initialdir=self.dir_salvar_excel,
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )

        if not excel_path:
            self.log("Operação cancelada pelo usuário.", "info")
            return

        # Atualizar o último diretório usado para salvar
        self.dir_salvar_excel = str(Path(excel_path).parent)
        self.salvar_configuracoes(silencioso=True)

        try:
            # Criar Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Ponto"

            # Cabeçalhos
            cabecalhos = [
                "FUNCIONÁRIO", "DIA", "MARCAÇÕES", "ENT. 1", "SAÍ. 1",
                "ENT. 2", "SAÍ. 2", "ENT. 3", "SAÍ. 3", "DURAÇÃO", "CH"
            ]
            for col, cabecalho in enumerate(cabecalhos, 1):
                cell = ws.cell(row=1, column=col, value=cabecalho)
                cell.font = cell.font.copy(bold=True)

            # Dados
            for row_idx, dia in enumerate(dados, 2):
                ws.cell(row=row_idx, column=1, value=dia.get("nome", ""))
                ws.cell(row=row_idx, column=2, value=dia.get("dia", ""))
                ws.cell(row=row_idx, column=3, value=dia.get("marcacoes", ""))
                ws.cell(row=row_idx, column=4, value=dia.get("ent1", ""))
                ws.cell(row=row_idx, column=5, value=dia.get("sai1", ""))
                ws.cell(row=row_idx, column=6, value=dia.get("ent2", ""))
                ws.cell(row=row_idx, column=7, value=dia.get("sai2", ""))
                ws.cell(row=row_idx, column=8, value=dia.get("ent3", ""))
                ws.cell(row=row_idx, column=9, value=dia.get("sai3", ""))
                ws.cell(row=row_idx, column=10, value=dia.get("duracao", ""))
                ws.cell(row=row_idx, column=11, value=dia.get("ch", ""))

            # Ajustar largura
            larguras = [30, 15, 25, 12, 12, 12, 12, 12, 12, 12, 10]
            for col, largura in enumerate(larguras, 1):
                ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = largura

            wb.save(excel_path)
            self.log(f"Planilha salva: {excel_path}", "success")

            # Mostrar mensagem de sucesso
            messagebox.showinfo(
                "Sucesso",
                f"Planilha Excel criada com sucesso!\n\n"
                f"Arquivo: {Path(excel_path).name}\n"
                f"Local: {Path(excel_path).parent}\n"
                f"Dias processados: {len(dados)}"
            )

        except Exception as e:
            self.log(f"Erro ao salvar Excel: {str(e)}", "error")
            messagebox.showerror("Erro", f"Não foi possível salvar a planilha:\n{str(e)}")

    def iniciar_processamento(self):
        """Inicia o processamento em thread separada."""
        if not self.pdf_paths:
            messagebox.showwarning("Atenção", "Selecione pelo menos um arquivo.")
            return

        if self.processamento_ativo:
            return

        # Confirmar processamento
        resposta = messagebox.askyesno(
            "Confirmar Processamento",
            f"Deseja processar {len(self.pdf_paths)} arquivo(s) PDF?\n\n"
            "O processo pode levar alguns segundos dependendo do tamanho dos arquivos."
        )

        if not resposta:
            return

        self.btn_processar.configure(state="disabled")
        self.log("--- Iniciando processamento ---", "processing")
        self.progress_var.set(0)

        # Executar em thread
        thread = threading.Thread(target=self.processar_pdfs_thread, daemon=True)
        thread.start()

    def _config_path(self):
        """Retorna o caminho do config.json em %APPDATA%\\ControlIDReader\\."""
        appdata = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        config_dir = appdata / "ControlIDReader"
        config_dir.mkdir(parents=True, exist_ok=True)
        return config_dir / "config.json"

    def carregar_configuracoes(self):
        """Carrega configurações salvas."""
        try:
            config_file = self._config_path()
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if content:
                        config = json.loads(content)
                        if config.get("dir_abrir_pdf") and Path(config["dir_abrir_pdf"]).exists():
                            self.dir_abrir_pdf = config["dir_abrir_pdf"]
                        if config.get("dir_salvar_excel") and Path(config["dir_salvar_excel"]).exists():
                            self.dir_salvar_excel = config["dir_salvar_excel"]
                        if config.get("tema") in ("dark", "light"):
                            self._tema_atual = config["tema"]
                            ctk.set_appearance_mode(self._tema_atual)
                            label = "🌙  Dark" if self._tema_atual == "dark" else "☀️  Light"
                            self.btn_tema.configure(text=label)
                            self._atualizar_cores_log()
        except Exception:
            pass  # Ignora erros de configuração — defaults já estão definidos no __init__

    def salvar_configuracoes(self, silencioso=False):
        """Salva configurações atuais."""
        try:
            config = {
                "dir_abrir_pdf": self.dir_abrir_pdf,
                "dir_salvar_excel": self.dir_salvar_excel,
                "tema": self._tema_atual,
                "data_ultima_execucao": datetime.now().isoformat()
            }
            config_file = self._config_path()
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            if not silencioso:
                self.log("Configurações salvas com sucesso.", "success")
                messagebox.showinfo("Sucesso", "Configurações salvas!")
        except Exception as e:
            self.log(f"Erro ao salvar configurações: {e}", "error")
            if not silencioso:
                messagebox.showerror("Erro", f"Não foi possível salvar configurações:\n{str(e)}")


def main():
    """Função principal."""
    root = ctk.CTk()
    app = PdfToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

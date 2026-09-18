import pdfplumber
from bs4 import BeautifulSoup
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import customtkinter as ctk
from datetime import datetime
import threading
import time

from control_id_reader import APP_VERSION
from control_id_reader.audit import analyze_inconsistencies, build_preview_summary
from control_id_reader.config_store import default_config_path, load_config, save_config
from control_id_reader.excel_writer import generate_suggested_filename, save_excel_file
from control_id_reader.parsers import (
    extract_employee_info_from_mhtml,
    extract_employee_info_from_text,
    extract_html_from_mhtml,
    extract_pdf_employees,
    extract_punch_rows_from_mhtml,
    extract_punch_rows_from_tables,
)
from control_id_reader.pdf_splitter import detect_employee_page_groups, write_employee_zip
from control_id_reader.privacy import (
    PrivacyError,
    anonymize_excel_file,
    restore_excel_file,
    suggest_anonymized_path,
    suggest_key_path,
    suggest_restored_path,
)
from control_id_reader.utils import (
    formatar_cpf,
)


class PdfToExcelApp:
    """Aplicativo moderno para converter folhas de ponto em Excel."""

    def __init__(self, root):
        self.root = root
        self.root.title(f"Control ID Reader {APP_VERSION} - Conversor de Ponto")
        self.root.geometry("700x780")
        self.root.resizable(True, True)
        self.root.minsize(700, 780)

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
        self.separacao_ativa = False
        self.privacidade_ativa = False
        self._elapsed_log_stop_event = None
        self._elapsed_log_interval = 5
        self._tema_atual = "dark"
        self._janela_ajuda = None

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

    def mostrar_ajuda(self):
        """Abre a ajuda sem interromper o uso da janela principal."""
        if self._janela_ajuda is not None and self._janela_ajuda.winfo_exists():
            self._janela_ajuda.lift()
            self._janela_ajuda.focus()
            return

        janela = ctk.CTkToplevel(self.root)
        self._janela_ajuda = janela
        janela.title("Ajuda - Control ID Reader")
        janela.minsize(520, 460)
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - 660) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - 620) // 2)
        janela.geometry(f"660x620+{x}+{y}")
        janela.transient(self.root)

        conteudo = ctk.CTkFrame(janela, fg_color="transparent")
        conteudo.pack(fill="both", expand=True, padx=20, pady=16)
        ctk.CTkLabel(
            conteudo, text="Ajuda", anchor="w",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(fill="x", pady=(0, 10))

        rolagem = ctk.CTkScrollableFrame(conteudo, fg_color="transparent")
        rolagem.pack(fill="both", expand=True)

        secoes = (
            ("Arquivos", (
                ("Adicionar Arquivos", "Inclui PDFs de espelho ou cartão de ponto e arquivos MHTML. É possível selecionar vários arquivos para a mesma planilha."),
                ("Remover Selecionados", "Retira da lista os arquivos marcados. Use Ctrl ou Shift para selecionar vários itens."),
                ("Limpar Tudo", "Esvazia a lista de seleção sem apagar os arquivos do computador."),
            )),
            ("Processamento", (
                ("Gerar Planilha Excel", "Lê os arquivos da lista e mostra uma pré-visualização antes de pedir onde salvar o Excel."),
                ("Pré-visualização", "Mostra funcionários, dias trabalhados, faltas e avisos. Gerar Excel confirma a exportação; Cancelar interrompe o salvamento."),
                ("Separar PDF por Funcionário", "Escolhe um PDF diretamente e cria um ZIP com um PDF por funcionário. Funciona independentemente da lista de arquivos."),
                ("Log, progresso e status", "Mostram as etapas, avisos, erros e andamento. Durante a separação, o log também informa o tempo decorrido."),
            )),
            ("Planilha gerada", (
                ("Ponto", "Reúne os registros diários. No cartão, PREVISTO indica a jornada e TOTAL NOTURNO aparece em coluna própria."),
                ("Inconsistências", "Lista avisos como dias previstos sem marcações e pares de entrada e saída incompletos. A aba aparece quando há avisos."),
                ("Resumo CH", "Agrupa dias trabalhados e faltados por código de horário contratual. Aparece quando o arquivo informa códigos CH."),
                ("Resumo", "Mostra, por funcionário, dias trabalhados, faltados, dias de trabalho totais e média de horas."),
                ("Espelho e cartão", "No espelho, o CH define o dia de trabalho. No cartão, vale a jornada PREVISTO; abonos e atestados identificados não contam como falta."),
            )),
            ("Privacidade", (
                ("Anonimizar Planilha", "Cria uma cópia do Excel com identificadores substituídos e uma chave .cidkey protegida por senha."),
                ("Restaurar Dados", "Recupera os identificadores usando a planilha anonimizada, a chave .cidkey e a senha."),
                ("Antes de compartilhar", "Confira textos livres, como justificativas do cartão, pois podem conter informações sensíveis que não são substituídas."),
            )),
            ("Preferências", (
                ("Tema", "Alterna entre os modos claro e escuro."),
                ("Salvar Configurações", "Guarda o tema e as últimas pastas usadas para abrir e salvar arquivos."),
            )),
        )

        for indice, (titulo, itens) in enumerate(secoes):
            if indice:
                ctk.CTkFrame(
                    rolagem, height=1, fg_color=("gray80", "gray30"),
                ).pack(fill="x", pady=(14, 12))
            ctk.CTkLabel(
                rolagem, text=titulo, anchor="w",
                font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(fill="x", pady=(0, 7))
            for nome, descricao in itens:
                ctk.CTkLabel(
                    rolagem, text=nome, anchor="w",
                    font=ctk.CTkFont(size=12, weight="bold"),
                ).pack(fill="x", pady=(6, 0))
                ctk.CTkLabel(
                    rolagem, text=descricao, anchor="w", justify="left",
                    wraplength=445, text_color=("gray35", "gray75"),
                    font=ctk.CTkFont(size=12),
                ).pack(fill="x", pady=(2, 0))

        ctk.CTkButton(
            conteudo, text="Fechar", width=100, command=janela.destroy,
        ).pack(anchor="e", pady=(12, 0))
        janela.bind("<Escape>", lambda _event: janela.destroy())
        janela.focus()

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
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

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

        self.btn_ajuda = ctk.CTkButton(
            header, text="Ajuda", width=72, height=30,
            fg_color=("gray80", "gray25"), text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray35"),
            command=self.mostrar_ajuda
        )
        self.btn_ajuda.grid(row=0, column=1, rowspan=2, sticky="e", padx=(0, 8))

        # Botão tema no canto superior direito
        self.btn_tema = ctk.CTkButton(
            header, text="🌙  Dark", width=100, height=30,
            fg_color=("gray80", "gray25"), text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray35"),
            command=self.alternar_tema
        )
        self.btn_tema.grid(row=0, column=2, rowspan=2, sticky="e")

        # Seções
        self.criar_secao_selecao(main_frame, 1)
        self.criar_secao_status(main_frame, 2)
        self.criar_secao_acoes(main_frame, 3)
    def criar_secao_selecao(self, parent, row):
        """Cria a seção para seleção de arquivos."""
        frame = ctk.CTkFrame(parent)
        frame.grid(row=row, column=0, sticky="nsew", pady=(0, 6))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ctk.CTkLabel(frame, text="Seleção de Arquivos",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).grid(row=0, column=0, sticky="w", padx=12, pady=(6, 3))

        # Lista de arquivos (tk.Listbox ainda não tem substituto CTk nativo)
        list_frame = ctk.CTkFrame(frame, fg_color="transparent")
        list_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 2))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.lista_arquivos = tk.Listbox(
            list_frame, height=20, selectmode=tk.EXTENDED,
            font=("Segoe UI", 9), relief="flat", borderwidth=0,
            activestyle="none", selectbackground="#1f6aa5",
            selectforeground="white"
        )
        self.lista_arquivos.grid(row=0, column=0, sticky="nsew")

        scrollbar = ctk.CTkScrollbar(list_frame, command=self.lista_arquivos.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.lista_arquivos.configure(yscrollcommand=scrollbar.set)

        # Botões
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="w", padx=12, pady=(2, 2))

        self.btn_adicionar = ctk.CTkButton(
            btn_frame, text="+ Adicionar Arquivos", width=160,
            fg_color="#27ae60", hover_color="#1e8449",
            command=self.adicionar_pdfs
        )
        self.btn_adicionar.grid(row=0, column=0, padx=(0, 6))

        self.btn_remover = ctk.CTkButton(
            btn_frame, text="✕ Remover Selecionados", width=170,
            fg_color="#e67e22", hover_color="#ca6f1e",
            command=self.remover_selecionados
        )
        self.btn_remover.grid(row=0, column=1, padx=(0, 6))

        self.btn_limpar = ctk.CTkButton(
            btn_frame, text="🗑  Limpar Tudo", width=130,
            fg_color="#c0392b", hover_color="#a93226",
            command=self.limpar_lista
        )
        self.btn_limpar.grid(row=0, column=2)

        self.info_arquivos = ctk.CTkLabel(
            frame, text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=11), text_color=("gray40", "gray60")
        )
        self.info_arquivos.grid(row=3, column=0, sticky="w", padx=12, pady=(0, 6))

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
            frame, height=20, wrap=tk.WORD,
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
        button_width = 220
        button_height = 40

        # Barra de progresso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(frame, variable=self.progress_var,
                                               height=10)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, columnspan=3, sticky="ew",
                               pady=(0, 6))

        self.status_final = ctk.CTkLabel(
            frame, text="",
            font=ctk.CTkFont(size=11), text_color=("gray40", "gray60")
        )
        self.status_final.grid(row=1, column=0, columnspan=3,
                               sticky="w", pady=(0, 8))

        # Botão principal
        self.btn_processar = ctk.CTkButton(
            frame, text="⚡  Gerar Planilha Excel",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=button_height, width=button_width,
            fg_color="#2980b9", hover_color="#2471a3",
            command=self.iniciar_processamento
        )
        self.btn_processar.grid(row=2, column=0, sticky="w")

        self.btn_anonimizar = ctk.CTkButton(
            frame, text="🔒  Anonimizar Planilha",
            height=button_height, width=button_width,
            fg_color="#16a085", hover_color="#138d75",
            command=self.anonimizar_planilha
        )
        self.btn_anonimizar.grid(row=2, column=1, padx=(10, 0), sticky="w")

        self.btn_separar_pdf = ctk.CTkButton(
            frame, text="✂️  Separar PDF por Funcionário",
            height=button_height, width=button_width,
            fg_color="#7d3c98", hover_color="#6c3483",
            command=self.separar_pdf_por_funcionario
        )
        self.btn_separar_pdf.grid(row=2, column=2, padx=(10, 0), sticky="w")

        self.btn_salvar_config = ctk.CTkButton(
            frame, text="💾  Salvar Configurações",
            height=button_height, width=button_width,
            fg_color=("gray75", "gray30"), text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray40"),
            command=self.salvar_configuracoes
        )
        self.btn_salvar_config.grid(row=3, column=0, sticky="w", pady=(8, 0))

        self.btn_restaurar = ctk.CTkButton(
            frame, text="🔓  Restaurar Dados",
            height=button_height, width=button_width,
            fg_color="#8e44ad", hover_color="#7d3c98",
            command=self.restaurar_planilha
        )
        self.btn_restaurar.grid(row=3, column=1, padx=(10, 0), sticky="w", pady=(8, 0))

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

    def _start_elapsed_log(self, mensagem, tipo="processing", intervalo=None):
        """Registra no log o tempo decorrido de uma operação longa."""
        self._stop_elapsed_log()
        intervalo = intervalo or self._elapsed_log_interval
        stop_event = threading.Event()
        self._elapsed_log_stop_event = stop_event
        inicio = time.monotonic()

        def registrar_tempo():
            while not stop_event.wait(intervalo):
                decorrido = int(time.monotonic() - inicio)
                self.log(f"{mensagem} em andamento há {self._formatar_tempo_decorrido(decorrido)}.", tipo)

        threading.Thread(target=registrar_tempo, daemon=True).start()

    def _stop_elapsed_log(self, mensagem_final=None, tipo="processing", status_final=None):
        if self._elapsed_log_stop_event:
            self._elapsed_log_stop_event.set()
            self._elapsed_log_stop_event = None
        if mensagem_final is not None:
            self.log(mensagem_final, tipo)
        if status_final is not None:
            self.root.after(0, lambda: self.status_final.configure(text=status_final))

    def _formatar_tempo_decorrido(self, segundos):
        minutos, seg = divmod(max(0, segundos), 60)
        if minutos:
            return f"{minutos}min {seg:02d}s"
        return f"{seg}s"

    def _set_progress(self, valor, status=None):
        """Atualiza barra e status a partir da thread principal."""
        def _aplicar():
            self.progress_bar.set(valor)
            if status is not None:
                self.status_final.configure(text=status)
        self.root.after(0, _aplicar)

    def _set_acoes_habilitadas(self, habilitado):
        """Habilita ou desabilita ações que não devem concorrer com processamento."""
        estado = "normal" if habilitado else "disabled"
        self.btn_adicionar.configure(state=estado)
        self.btn_remover.configure(state=estado)
        self.btn_limpar.configure(state=estado)
        self.btn_processar.configure(state=estado)
        self.btn_salvar_config.configure(state=estado)
        self.btn_separar_pdf.configure(state=estado)
        self.btn_anonimizar.configure(state=estado)
        self.btn_restaurar.configure(state=estado)
        self.lista_arquivos.configure(state=estado)

    def _finalizar_separacao_ui(self, status=None):
        """Restaura a UI depois da separação de PDF."""
        def _aplicar():
            self.separacao_ativa = False
            self._set_acoes_habilitadas(True)
            if status is not None:
                self.status_final.configure(text=status)
        self.root.after(0, _aplicar)

    def _finalizar_processamento(self, todos_dias, total_pdfs, total_dias, todos_funcionarios):
        """Chamado na thread principal ao fim do processamento para abrir o filedialog."""
        self.processamento_ativo = False
        if todos_dias:
            inconsistencias = analyze_inconsistencies(todos_dias, todos_funcionarios)
            if not self._mostrar_pre_visualizacao(todos_dias, todos_funcionarios, inconsistencias, total_pdfs):
                self.log("Exportação cancelada na pré-visualização.", "info")
                self.status_final.configure(text="Processamento concluído; exportação cancelada.")
                self._set_acoes_habilitadas(True)
                return

            if self.salvar_excel(todos_dias, todos_funcionarios, inconsistencias):
                self.log(f"\n✅ Processamento concluído!", "success")
                self.log(f"Total: {total_pdfs} arquivo(s), {total_dias} dia(s) registrado(s)", "success")
                if inconsistencias:
                    self.log(f"Inconsistências registradas: {len(inconsistencias)}", "processing")
                self.status_final.configure(
                    text=f"Concluído: {total_pdfs} arquivo(s), {total_dias} dia(s)."
                )
            else:
                self.status_final.configure(text="Processamento concluído; salvamento cancelado.")
            self._set_acoes_habilitadas(True)
        else:
            self.log("\n⚠️ Nenhum dado foi extraído dos arquivos.", "error")
            self.status_final.configure(text="Nenhum dado foi extraído.")
            self._set_acoes_habilitadas(True)

    def _mostrar_pre_visualizacao(self, dados, todos_funcionarios, inconsistencias, total_pdfs):
        """Mostra uma janela de pré-visualização e retorna True se o usuário confirmar."""
        resumo = build_preview_summary(dados, todos_funcionarios, inconsistencias)
        resultado = {"confirmado": False}

        janela = ctk.CTkToplevel(self.root)
        janela.title("Pré-visualização da exportação")
        janela.geometry("760x560")
        janela.minsize(680, 480)
        janela.transient(self.root)
        janela.grab_set()

        frame = ctk.CTkFrame(janela)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ctk.CTkLabel(
            frame,
            text="Pré-visualização",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))

        texto_resumo = (
            f"Arquivos: {total_pdfs}    "
            f"Funcionários: {resumo['total_funcionarios']}    "
            f"Dias: {resumo['total_dias']}    "
            f"Trabalhados: {resumo['total_dias_trabalhados']}    "
            f"Faltados: {resumo['total_dias_faltados']}    "
            f"Dias de trabalho: {resumo['total_dias_trabalho']}    "
            f"Inconsistências: {resumo['total_inconsistencias']}"
        )
        ctk.CTkLabel(
            frame,
            text=texto_resumo,
            font=ctk.CTkFont(size=12),
            text_color=("gray30", "gray75"),
        ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 8))

        preview = ctk.CTkTextbox(frame, wrap="word")
        preview.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
        preview.insert("end", self._texto_pre_visualizacao(resumo, inconsistencias))
        preview.configure(state="disabled")

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=3, column=0, sticky="e", padx=12, pady=(0, 12))

        def cancelar():
            resultado["confirmado"] = False
            janela.destroy()

        def confirmar():
            resultado["confirmado"] = True
            janela.destroy()

        ctk.CTkButton(
            botoes,
            text="Cancelar",
            width=120,
            fg_color=("gray75", "gray30"),
            text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray40"),
            command=cancelar,
        ).grid(row=0, column=0, padx=(0, 8))

        ctk.CTkButton(
            botoes,
            text="Gerar Excel",
            width=140,
            fg_color="#27ae60",
            hover_color="#1e8449",
            command=confirmar,
        ).grid(row=0, column=1)

        janela.protocol("WM_DELETE_WINDOW", cancelar)
        janela.wait_window()
        return resultado["confirmado"]

    def _texto_pre_visualizacao(self, resumo, inconsistencias):
        linhas = ["FUNCIONÁRIOS", ""]
        for funcionario in resumo["funcionarios"]:
            cpf = funcionario["cpf"] or "CPF não identificado"
            linhas.append(
                f"- {funcionario['nome'] or 'Sem nome'} | {cpf} | "
                f"{funcionario['dias']} dia(s), "
                f"{funcionario['dias_trabalhados']} trabalhado(s), "
                f"{funcionario['dias_faltados']} faltado(s), "
                f"{funcionario['dias_trabalho_totais']} dia(s) de trabalho"
            )

        linhas.extend(["", "INCONSISTÊNCIAS", ""])
        if not inconsistencias:
            linhas.append("Nenhuma inconsistência encontrada.")
        else:
            por_severidade = resumo["inconsistencias_por_severidade"]
            resumo_severidade = ", ".join(
                f"{sev}: {qtd}" for sev, qtd in sorted(por_severidade.items())
            )
            linhas.append(resumo_severidade)
            linhas.append("")
            for issue in inconsistencias[:80]:
                trecho_dia = f" | {issue['dia']}" if issue.get("dia") else ""
                linhas.append(
                    f"[{issue['severidade']}] {issue.get('funcionario') or 'Sem funcionário'}"
                    f"{trecho_dia} | {issue['campo']}: {issue['mensagem']}"
                )
            if len(inconsistencias) > 80:
                linhas.append(f"... mais {len(inconsistencias) - 80} inconsistência(s).")

        return "\n".join(linhas)

    def extrair_info_funcionario(self, texto):
        """Extrai informações do funcionário a partir do texto de um arquivo.
        Retorna (dict, motivo_falha). motivo_falha é None em caso de sucesso."""
        return extract_employee_info_from_text(texto)

    def extrair_dados_ponto(self, tabelas):
        """Extrai dados de ponto e horários contratuais de uma lista de tabelas.
        Retorna (dias, horarios_contratuais).

        Suporta múltiplos formatos de PDF — a detecção é semântica:
        - ent2/sai2 podem estar mesclados em [4] ou divididos por pdfplumber
        - CH (5 dígitos) e duração (HH:MM) são localizados por valor, não por índice fixo
        """
        return extract_punch_rows_from_tables(tabelas)

    # ──────────────────────────────────────────────
    # Extração MHTML
    # ──────────────────────────────────────────────

    def extrair_html_de_mhtml(self, caminho):
        """Abre um arquivo .mhtml e retorna o conteúdo HTML interno como string."""
        return extract_html_from_mhtml(caminho)

    def extrair_info_funcionario_mhtml(self, soup):
        """Extrai dados do funcionário a partir do HTML parseado.
        Retorna (dict, motivo_falha). Células têm formato 'CAMPO:Valor'."""
        return extract_employee_info_from_mhtml(soup)

    def extrair_dados_ponto_mhtml(self, soup):
        """Extrai registros de ponto da tabela HTML.
        Cada linha-mestre tem 17 células: [0]=data, [3]=ENT1, [4]=SAÍ1,
        [5]=ENT2, [6]=SAÍ2. Linhas de detalhe (3 células) são ignoradas."""
        return extract_punch_rows_from_mhtml(soup)

    # ──────────────────────────────────────────────

    def extrair_funcionarios_de_pdf(self, pdf):
        """Processa um PDF com um ou mais funcionários, página a página.
        Retorna lista de tuplas (info_funcionario, dias)."""
        return extract_pdf_employees(pdf, logger=self.log)

    def processar_pdfs_thread(self):
        """Thread principal de processamento."""
        self.processamento_ativo = True
        total_pdfs = len(self.pdf_paths)
        total_dias = 0
        todos_dias = []
        todos_funcionarios = {}  # chave: CPF → info_funcionario completo

        try:
            self.log(f"Iniciando processamento de {total_pdfs} arquivo(s)...", "processing")

            for idx, pdf_path in enumerate(self.pdf_paths):
                nome_arquivo = Path(pdf_path).name
                self.log(f"Processando: {nome_arquivo}", "processing")

                try:
                    if not Path(pdf_path).exists():
                        self.log(f"  ✗ Arquivo não encontrado: {nome_arquivo}", "error")
                        continue

                    ext = Path(pdf_path).suffix.lower()

                    if ext == ".mhtml":
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

                    elif ext == ".pdf":
                        try:
                            pdf_aberto = pdfplumber.open(pdf_path)
                        except Exception:
                            self.log(f"  ✗ Não foi possível abrir '{nome_arquivo}'. "
                                     f"O arquivo pode estar protegido por senha ou corrompido.", "error")
                            continue

                        with pdf_aberto as pdf:
                            funcionarios_no_pdf = self.extrair_funcionarios_de_pdf(pdf)

                        if not funcionarios_no_pdf:
                            self.log(f"  ✗ '{nome_arquivo}' não contém texto extraível ou "
                                     f"nenhum funcionário foi identificado.", "error")
                            continue

                        for info_funcionario, dias_do_pdf in funcionarios_no_pdf:
                            cpf = info_funcionario.get("cpf", info_funcionario.get("nome", ""))
                            for dia in dias_do_pdf:
                                dia["nome"]    = info_funcionario.get("nome", "")
                                dia["empresa"] = info_funcionario.get("empresa", "")
                                dia["cpf"]     = formatar_cpf(info_funcionario.get("cpf", ""))

                            if cpf not in todos_funcionarios:
                                todos_funcionarios[cpf] = info_funcionario
                            todos_funcionarios[cpf].setdefault("_dias", []).extend(dias_do_pdf)

                            todos_dias.extend(dias_do_pdf)
                            total_dias += len(dias_do_pdf)
                            self.log(f"  ✓ {info_funcionario.get('nome', cpf)}: "
                                     f"{len(dias_do_pdf)} dia(s) extraído(s)", "success")
                        continue  # pula o bloco de associação abaixo (já feito acima)

                    else:
                        self.log(f"  ✗ Formato não suportado: {ext or 'sem extensão'}", "error")
                        continue

                    # Associar dados básicos a cada dia e registrar funcionário completo (MHTML)
                    cpf = info_funcionario.get("cpf", info_funcionario.get("nome", ""))
                    for dia in dias_do_pdf:
                        dia["nome"]    = info_funcionario.get("nome", "")
                        dia["empresa"] = info_funcionario.get("empresa", "")
                        dia["cpf"]     = formatar_cpf(info_funcionario.get("cpf", ""))

                    # Acumula dias do mesmo funcionário somando os que já existem
                    if cpf not in todos_funcionarios:
                        todos_funcionarios[cpf] = info_funcionario
                    todos_funcionarios[cpf].setdefault("_dias", []).extend(dias_do_pdf)

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
                finally:
                    progresso = (idx + 1) / total_pdfs if total_pdfs else 0
                    self.root.after(0, lambda p=progresso: self.progress_bar.set(p))

        except Exception as e:
            self.log(f"Erro crítico no processamento: {str(e)}", "error")
        finally:
            self.root.after(0, lambda: self._finalizar_processamento(
                todos_dias, total_pdfs, total_dias, todos_funcionarios))

    def salvar_excel(self, dados, todos_funcionarios, inconsistencias=None):
        """Salva os dados em um arquivo Excel com cabeçalho, aba de resumo e nome inteligente."""

        # ── Nome do arquivo inteligente ─────────────────────────────────────
        nome_sugerido = self._gerar_nome_arquivo(dados, todos_funcionarios)

        excel_path = filedialog.asksaveasfilename(
            title="Salvar planilha Excel",
            defaultextension=".xlsx",
            initialfile=nome_sugerido,
            initialdir=self.dir_salvar_excel,
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )

        if not excel_path:
            self.log("Operação cancelada pelo usuário.", "info")
            return False

        self.dir_salvar_excel = str(Path(excel_path).parent)
        self.salvar_configuracoes(silencioso=True)

        try:
            save_excel_file(excel_path, dados, todos_funcionarios, inconsistencias=inconsistencias)
            self.log(f"Planilha salva: {excel_path}", "success")
            total_inconsistencias = len(inconsistencias or [])

            messagebox.showinfo(
                "Sucesso",
                f"Planilha Excel criada com sucesso!\n\n"
                f"Arquivo: {Path(excel_path).name}\n"
                f"Local: {Path(excel_path).parent}\n"
                f"Funcionários: {len(todos_funcionarios)}\n"
                f"Dias processados: {len(dados)}\n"
                f"Inconsistências: {total_inconsistencias}"
            )
            return True

        except Exception as e:
            self.log(f"Erro ao salvar Excel: {str(e)}", "error")
            messagebox.showerror("Erro", f"Não foi possível salvar a planilha:\n{str(e)}")
            return False

    def _gerar_nome_arquivo(self, dados, todos_funcionarios):
        """Gera o nome sugerido para o arquivo Excel com base no cenário."""
        return generate_suggested_filename(dados, todos_funcionarios, len(self.pdf_paths))

    def anonimizar_planilha(self):
        """Cria uma cópia anonimizada da planilha e uma chave criptografada."""
        if self.privacidade_ativa:
            return

        excel_path = filedialog.askopenfilename(
            title="Selecionar planilha para anonimizar",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialdir=self.dir_salvar_excel
        )
        if not excel_path:
            return

        output_path = filedialog.asksaveasfilename(
            title="Salvar planilha anonimizada",
            defaultextension=".xlsx",
            initialfile=suggest_anonymized_path(excel_path).name,
            initialdir=str(Path(excel_path).parent),
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if not output_path:
            return

        key_path = filedialog.asksaveasfilename(
            title="Salvar chave de restauração",
            defaultextension=".cidkey",
            initialfile=suggest_key_path(output_path).name,
            initialdir=str(Path(output_path).parent),
            filetypes=[("Chave Control ID Reader", "*.cidkey")]
        )
        if not key_path:
            return

        senha = self._pedir_senha_privacidade(confirmar=True)
        if not senha:
            return

        self._executar_privacidade_thread(
            "Anonimizando planilha...",
            anonymize_excel_file,
            (excel_path, output_path, key_path, senha),
            lambda result: (
                "Planilha anonimizada com sucesso!",
                (
                    f"Planilha anonimizada com sucesso!\n\n"
                    f"Arquivo: {Path(result['output_path']).name}\n"
                    f"Chave: {Path(result['key_path']).name}\n"
                    f"Valores únicos protegidos: {result['unique_values']}\n\n"
                    "Guarde a chave e a senha em segurança. Sem elas não dá para restaurar os dados reais."
                )
            ),
        )

    def restaurar_planilha(self):
        """Restaura uma planilha anonimizada usando a chave criptografada."""
        if self.privacidade_ativa:
            return

        excel_path = filedialog.askopenfilename(
            title="Selecionar planilha anonimizada",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialdir=self.dir_salvar_excel
        )
        if not excel_path:
            return

        key_path = filedialog.askopenfilename(
            title="Selecionar chave de restauração",
            filetypes=[("Chave Control ID Reader", "*.cidkey")],
            initialdir=str(Path(excel_path).parent)
        )
        if not key_path:
            return

        output_path = filedialog.asksaveasfilename(
            title="Salvar planilha restaurada",
            defaultextension=".xlsx",
            initialfile=suggest_restored_path(excel_path).name,
            initialdir=str(Path(excel_path).parent),
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if not output_path:
            return

        senha = self._pedir_senha_privacidade(confirmar=False)
        if not senha:
            return

        self._executar_privacidade_thread(
            "Restaurando dados da planilha...",
            restore_excel_file,
            (excel_path, output_path, key_path, senha),
            lambda result: (
                "Planilha restaurada com sucesso!",
                (
                    f"Planilha restaurada com sucesso!\n\n"
                    f"Arquivo: {Path(result['output_path']).name}\n"
                    f"Células/textos restaurados: {result['restored']}"
                )
            ),
        )

    def _pedir_senha_privacidade(self, confirmar=False):
        senha = simpledialog.askstring(
            "Senha da chave",
            "Digite uma senha para proteger a chave:" if confirmar else "Digite a senha da chave:",
            show="*",
            parent=self.root,
        )
        if senha is None:
            return None
        if not senha:
            messagebox.showwarning("Atenção", "A senha não pode ficar vazia.")
            return None

        if confirmar:
            confirmacao = simpledialog.askstring(
                "Confirmar senha",
                "Digite a senha novamente:",
                show="*",
                parent=self.root,
            )
            if confirmacao is None:
                return None
            if senha != confirmacao:
                messagebox.showwarning("Atenção", "As senhas não conferem.")
                return None

        return senha

    def _executar_privacidade_thread(self, status, funcao, args, mensagem_sucesso):
        self.privacidade_ativa = True
        self._set_acoes_habilitadas(False)
        self.progress_var.set(0)
        self.status_final.configure(text=status)
        self.log(status, "processing")

        def worker():
            try:
                resultado = funcao(*args)
                titulo, mensagem = mensagem_sucesso(resultado)

                def finalizar_sucesso():
                    self.progress_bar.set(1)
                    self.status_final.configure(text=titulo)
                    self.log(titulo, "success")
                    messagebox.showinfo("Concluído", mensagem)
                    self.privacidade_ativa = False
                    self._set_acoes_habilitadas(True)

                self.root.after(0, finalizar_sucesso)

            except PrivacyError as e:
                mensagem = str(e)
                self.root.after(0, lambda msg=mensagem: self._finalizar_privacidade_erro(msg))
            except PermissionError:
                self.root.after(0, lambda: self._finalizar_privacidade_erro(
                    "Sem permissão para ler ou salvar os arquivos. Verifique se a planilha está aberta."
                ))
            except Exception as e:
                mensagem = f"Erro inesperado: {str(e)}"
                self.root.after(0, lambda msg=mensagem: self._finalizar_privacidade_erro(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _finalizar_privacidade_erro(self, mensagem):
        self.progress_bar.set(0)
        self.status_final.configure(text="Operação de privacidade cancelada ou concluída com erro.")
        self.log(f"✗ {mensagem}", "error")
        messagebox.showerror("Erro", mensagem)
        self.privacidade_ativa = False
        self._set_acoes_habilitadas(True)

    def separar_pdf_por_funcionario(self):
        """Abre um PDF com múltiplos funcionários e gera um ZIP com um PDF por funcionário."""
        if self.separacao_ativa:
            return

        pdf_path = filedialog.askopenfilename(
            title="Selecionar PDF com múltiplos funcionários",
            filetypes=[("Arquivos PDF", "*.pdf")],
            initialdir=self.dir_abrir_pdf
        )
        if not pdf_path:
            return

        self.dir_abrir_pdf = str(Path(pdf_path).parent)
        self.salvar_configuracoes(silencioso=True)

        zip_path = filedialog.asksaveasfilename(
            title="Salvar ZIP com PDFs separados",
            defaultextension=".zip",
            initialfile=f"{Path(pdf_path).stem}_separados.zip",
            initialdir=self.dir_salvar_excel,
            filetypes=[("Arquivo ZIP", "*.zip")]
        )
        if not zip_path:
            return

        self.separacao_ativa = True
        self._set_acoes_habilitadas(False)
        self.progress_var.set(0)
        self.status_final.configure(text="Separando PDF: detectando funcionários...")
        self.log("Iniciando separação de PDF por funcionário...", "processing")
        self._start_elapsed_log("Separação de PDF")
        t = threading.Thread(
            target=self._separar_pdf_worker,
            args=(pdf_path, zip_path),
            daemon=True
        )
        t.start()

    def _separar_pdf_worker(self, pdf_path, zip_path):
        """Thread que separa o PDF e gera o ZIP."""
        try:
            self._set_progress(0.03, "Separando PDF: detectando funcionários...")

            def registrar_arquivo(arquivo, index, total):
                progresso = 0.15 + (0.85 * index / total)
                self._set_progress(
                    progresso,
                    f"Separando PDF: {index}/{total} funcionário(s) exportado(s)..."
                )
                self.log(
                    f"  ✓ {arquivo['nome']} → {arquivo['arquivo']} "
                    f"({arquivo['paginas']} página(s))",
                    "success",
                )

            grupos = detect_employee_page_groups(pdf_path)
            self._set_progress(0.15, f"Separando PDF: {len(grupos)} funcionário(s) detectado(s).")

            if not grupos:
                self._stop_elapsed_log(
                    "Separação cancelada: nenhum funcionário identificado.",
                    "error",
                    status_final="Separação cancelada: nenhum funcionário identificado.",
                )
                self.root.after(0, lambda: (
                    self.log("⚠️ Nenhum funcionário identificado no PDF.", "error"),
                    messagebox.showwarning("Atenção", "Nenhum funcionário identificado no PDF.")
                ))
                self._finalizar_separacao_ui("Separação cancelada: nenhum funcionário identificado.")
                return

            if len(grupos) == 1:
                self._stop_elapsed_log(
                    "Separação cancelada: PDF contém apenas um funcionário.",
                    "error",
                    status_final="Separação cancelada: PDF contém apenas um funcionário.",
                )
                self.root.after(0, lambda: (
                    self.log("⚠️ O PDF contém apenas um funcionário — separação desnecessária.", "error"),
                    messagebox.showwarning("Atenção", "O PDF contém apenas um funcionário.")
                ))
                self._finalizar_separacao_ui("Separação cancelada: PDF contém apenas um funcionário.")
                return

            self.log(f"  {len(grupos)} funcionário(s) detectado(s). Gerando PDFs...", "processing")
            write_employee_zip(pdf_path, zip_path, grupos, progress_callback=registrar_arquivo)

            self.root.after(0, lambda: (
                self.log(f"\n✅ ZIP gerado com sucesso: {zip_path}", "success"),
                messagebox.showinfo(
                    "Concluído",
                    f"PDFs separados com sucesso!\n\n"
                    f"Funcionários: {len(grupos)}\n"
                    f"Arquivo: {Path(zip_path).name}\n"
                    f"Local: {Path(zip_path).parent}"
                )
            ))
            self._set_progress(1, "Separação concluída.")
            self._stop_elapsed_log(
                "Separação concluída.",
                "success",
                status_final="Separação concluída.",
            )
            self._finalizar_separacao_ui("Separação concluída.")

        except PermissionError:
            self._stop_elapsed_log(
                "Erro na separação: sem permissão.",
                "error",
                status_final="Erro na separação: sem permissão.",
            )
            self.root.after(0, lambda: (
                self.log("✗ Sem permissão para ler o PDF ou salvar o ZIP.", "error"),
                messagebox.showerror("Erro", "Sem permissão para acessar o arquivo.")
            ))
            self._finalizar_separacao_ui("Erro na separação: sem permissão.")
        except Exception as e:
            self._stop_elapsed_log(
                "Erro ao separar PDF.",
                "error",
                status_final="Erro ao separar PDF.",
            )
            self.root.after(0, lambda: (
                self.log(f"✗ Erro ao separar PDF: {str(e)}", "error"),
                messagebox.showerror("Erro", f"Erro ao separar PDF:\n{str(e)}")
            ))
            self._finalizar_separacao_ui("Erro ao separar PDF.")

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
            f"Deseja processar {len(self.pdf_paths)} arquivo(s) PDF/MHTML?\n\n"
            "O processo pode levar alguns segundos dependendo do tamanho dos arquivos."
        )

        if not resposta:
            return

        self.processamento_ativo = True
        self._set_acoes_habilitadas(False)
        self.log("--- Iniciando processamento ---", "processing")
        self.status_final.configure(text="Processando arquivos...")
        self.progress_var.set(0)

        # Executar em thread
        thread = threading.Thread(target=self.processar_pdfs_thread, daemon=True)
        thread.start()

    def _config_path(self):
        """Retorna o caminho do config.json em %APPDATA%\\ControlIDReader\\."""
        return default_config_path()

    def carregar_configuracoes(self):
        """Carrega configurações salvas."""
        try:
            config = load_config(self._config_path())
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
            save_config(config, self._config_path())
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

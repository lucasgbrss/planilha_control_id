import pdfplumber
from pathlib import Path
from openpyxl import Workbook
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
import threading
import json
import os


class PdfToExcelApp:
    """Aplicativo moderno para converter folhas de ponto PDF em Excel."""

    def __init__(self, root):
        self.root = root
        self.root.title("Control ID Reader - Conversor de Ponto")
        self.root.geometry("780x620")
        self.root.resizable(False, False)

        # Aplicar ícone da janela
        self._aplicar_icone()

        # Centralizar janela na tela
        self.centralizar_janela()

        # Variáveis de estado
        self.pdf_paths = []
        self.processamento_ativo = False

        # Diretórios padrão (serão sobrescritos pelas configurações salvas)
        self.dir_abrir_pdf = str(Path.home() / "Downloads")
        self.dir_salvar_excel = str(Path.home() / "Documents")

        # Configurar estilo moderno
        self.configurar_estilo()

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

    def configurar_estilo(self):
        """Aplica um estilo mais moderno à interface."""
        style = ttk.Style()
        style.theme_use('clam')

        # Cores modernas
        style.configure('TFrame', background='white')
        style.configure('TLabel', background='white', foreground='#333')
        style.configure('TButton', background='#4a90d9', foreground='white',
                       padding=6, font=('Segoe UI', 9))
        style.map('TButton', background=[('active', '#3a7bc8')])

        # Botão verde — adicionar
        style.configure('Adicionar.TButton', background='#27ae60', foreground='white',
                        padding=6, font=('Segoe UI', 9))
        style.map('Adicionar.TButton', background=[('active', '#1e8449')])

        # Botão laranja — remover selecionados
        style.configure('Remover.TButton', background='#e67e22', foreground='white',
                        padding=6, font=('Segoe UI', 9))
        style.map('Remover.TButton', background=[('active', '#ca6f1e')])

        # Botão vermelho — limpar tudo
        style.configure('Limpar.TButton', background='#c0392b', foreground='white',
                        padding=6, font=('Segoe UI', 9))
        style.map('Limpar.TButton', background=[('active', '#a93226')])

        # Botão azul — ação principal
        style.configure('Principal.TButton', background='#2980b9', foreground='white',
                        padding=6, font=('Segoe UI', 9, 'bold'))
        style.map('Principal.TButton', background=[('active', '#2471a3')])

        # Botão cinza — ação secundária
        style.configure('Secundario.TButton', background='#7f8c8d', foreground='white',
                        padding=6, font=('Segoe UI', 9))
        style.map('Secundario.TButton', background=[('active', '#717d7e')])

        style.configure('Header.TLabel', font=('Segoe UI', 14, 'bold'),
                       background='white', foreground='#2c3e50')
        style.configure('Status.TLabel', font=('Segoe UI', 9), foreground='#666')
        style.configure('Success.TLabel', font=('Segoe UI', 9), foreground='#27ae60')
        style.configure('Error.TLabel', font=('Segoe UI', 9), foreground='#c0392b')
        style.configure('LabelFrame', background='white')
        style.configure('LabelFrame.Label', background='white', foreground='#2c3e50')

        style.configure('Treeview', rowheight=25, font=('Segoe UI', 9))
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))

    def criar_interface(self):
        """Cria todos os componentes da interface."""
        # Permitir que o frame principal expanda junto com a janela
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Container principal - centralizado
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        main_frame.columnconfigure(0, weight=1)

        # Título
        ttk.Label(main_frame, text="Control ID Reader", style='Header.TLabel')\
            .grid(row=0, column=0, pady=(0, 5))
        ttk.Label(main_frame, text="Conversor de Folhas de Ponto para Excel", foreground='#666')\
            .grid(row=1, column=0, pady=(0, 15))

        # Seção de seleção de arquivos
        self.criar_secao_selecao(main_frame, 2)

        # Seção de status/console
        self.criar_secao_status(main_frame, 3)

        # Seção de ações
        self.criar_secao_acoes(main_frame, 4)
    def criar_secao_selecao(self, parent, row):
        """Cria a seção para seleção de arquivos PDF."""
        frame = ttk.LabelFrame(parent, text=" Seleção de Arquivos ", padding="12")
        frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        # Lista de arquivos
        list_frame = ttk.Frame(frame)
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        list_frame.columnconfigure(0, weight=1)

        self.lista_arquivos = tk.Listbox(list_frame, height=5, selectmode=tk.EXTENDED,
                                        font=('Segoe UI', 9))
        self.lista_arquivos.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                  command=self.lista_arquivos.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.lista_arquivos.configure(yscrollcommand=scrollbar.set)

        # Botões de controle
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        ttk.Button(btn_frame, text="+ Adicionar PDFs",
                  command=self.adicionar_pdfs, style='Adicionar.TButton').grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="❌ Remover Selecionados",
                  command=self.remover_selecionados, style='Remover.TButton').grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="🗑️ Limpar Tudo",
                  command=self.limpar_lista, style='Limpar.TButton').grid(row=0, column=2, padx=5)

        # Info
        self.info_arquivos = ttk.Label(frame, text="Nenhum arquivo selecionado",
                                       style='Status.TLabel')
        self.info_arquivos.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))

    def criar_secao_status(self, parent, row):
        """Cria a seção de log/status."""
        frame = ttk.LabelFrame(parent, text=" Log de Processamento ", padding="10")
        frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(frame, height=8, wrap=tk.WORD,
                                                   font=('Consolas', 8))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Configurar cores de log
        self.log_text.tag_config('info', foreground='#333')
        self.log_text.tag_config('success', foreground='#27ae60')
        self.log_text.tag_config('error', foreground='#c0392b')
        self.log_text.tag_config('processing', foreground='#f39c12')

    def criar_secao_acoes(self, parent, row):
        """Cria a seção de botões de ação."""
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        frame.columnconfigure(0, weight=1)

        # Barra de progresso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var,
                                           maximum=100, mode='determinate')
        self.progress_bar.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E),
                             pady=(0, 5))

        # Botão de processar
        btn_processar = ttk.Button(frame, text="Gerar Planilha Excel",
                                   command=self.iniciar_processamento,
                                   style='Principal.TButton')
        btn_processar.grid(row=1, column=0, sticky=tk.W)

        self.btn_processar = btn_processar

        # Botão de salvar configurações
        ttk.Button(frame, text="Salvar Configurações",
                  command=self.salvar_configuracoes, style='Secundario.TButton').grid(row=1, column=1, padx=(10, 0))

        self.status_final = ttk.Label(frame, text="", style='Status.TLabel')
        self.status_final.grid(row=2, column=0, columnspan=3, pady=(5, 0))

    def adicionar_pdfs(self):
        """Abre dialog para selecionar arquivos PDF."""
        arquivo_types = [("Arquivos PDF", "*.pdf")]
        files = filedialog.askopenfilenames(
            title="Selecionar Arquivos PDF",
            filetypes=arquivo_types,
            initialdir=self.dir_abrir_pdf
        )

        if files:
            for f in files:
                if f not in self.pdf_paths:
                    self.pdf_paths.append(f)
                    self.lista_arquivos.insert(tk.END, Path(f).name)
            self.info_arquivos.config(text=f"{len(self.pdf_paths)} arquivo(s) selecionado(s)")

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

        self.info_arquivos.config(text=f"{len(self.pdf_paths)} arquivo(s) selecionado(s)")

    def limpar_lista(self):
        """Limpa todos os arquivos da lista."""
        self.pdf_paths.clear()
        self.lista_arquivos.delete(0, tk.END)
        self.info_arquivos.config(text="Nenhum arquivo selecionado")

    def log(self, mensagem, tipo='info'):
        """Adiciona mensagem ao log de forma thread-safe via root.after."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        def _inserir():
            self.log_text.insert(tk.END, f"[{timestamp}] {mensagem}\n", tipo)
            self.log_text.see(tk.END)
        self.root.after(0, _inserir)

    def _atualizar_progresso(self, valor):
        """Atualiza a barra de progresso de forma thread-safe."""
        self.root.after(0, lambda: self.progress_var.set(valor))

    def _finalizar_processamento(self, todos_dias, total_pdfs, total_dias):
        """Chamado na thread principal ao fim do processamento para abrir o filedialog."""
        self.processamento_ativo = False
        self.btn_processar.config(state='normal')
        if todos_dias:
            self.salvar_excel(todos_dias)
            self.log(f"\n✅ Processamento concluído!", "success")
            self.log(f"Total: {total_pdfs} PDF(s), {total_dias} dia(s) registrado(s)", "success")
        else:
            self.log("\n⚠️ Nenhum dado foi extraído dos PDFs.", "error")

    def extrair_info_funcionario(self, texto):
        """Extrai informações do funcionário a partir do texto de um PDF.
        Retorna (dict, motivo_falha). motivo_falha é None em caso de sucesso."""
        info = {}

        if not texto or not texto.strip():
            return info, "texto do PDF está vazio"

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
                motivo = "nenhum campo reconhecido — formato do PDF pode ser diferente do esperado"
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

    def processar_pdfs_thread(self):
        """Thread principal de processamento."""
        # Fix 7: flag ativado antes do try para garantir consistência com o botão desabilitado
        self.processamento_ativo = True
        total_pdfs = len(self.pdf_paths)
        total_dias = 0
        todos_dias = []

        try:
            self.log(f"Iniciando processamento de {total_pdfs} arquivo(s) PDF...", "processing")

            for idx, pdf_path in enumerate(self.pdf_paths):
                nome_arquivo = Path(pdf_path).name
                self.log(f"Processando: {nome_arquivo}", "processing")

                try:
                    # Fix 3: verificar se o arquivo ainda existe antes de abrir
                    if not Path(pdf_path).exists():
                        self.log(f"  ✗ Arquivo não encontrado: {nome_arquivo}", "error")
                        continue

                    texto_do_pdf = ""
                    tabelas_do_pdf = []

                    try:
                        pdf_aberto = pdfplumber.open(pdf_path)
                    except Exception:
                        # Fix 4: PDF protegido por senha ou corrompido
                        self.log(f"  ✗ Não foi possível abrir '{nome_arquivo}'. "
                                 f"O arquivo pode estar protegido por senha ou corrompido.", "error")
                        continue

                    with pdf_aberto as pdf:
                        # Fix 4: detectar PDF sem texto (possível senha ou só imagens)
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

                    # Extrair informações do funcionário
                    info_funcionario, motivo_falha = self.extrair_info_funcionario(texto_do_pdf)

                    if not info_funcionario.get("nome"):
                        # Fix 6: log com motivo específico da falha
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

                # Fix 2: atualizar progresso via root.after
                self._atualizar_progresso(((idx + 1) / total_pdfs) * 100)

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
            messagebox.showwarning("Atenção", "Selecione pelo menos um arquivo PDF.")
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

        self.btn_processar.config(state='disabled')
        self.log("--- Iniciando processamento ---", "processing")
        self.progress_var.set(0)

        # Executar em thread
        thread = threading.Thread(target=self.processar_pdfs_thread, daemon=True)
        thread.start()

    def carregar_configuracoes(self):
        """Carrega configurações salvas."""
        try:
            config_file = Path(__file__).parent / "config.json"
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if content:
                        config = json.loads(content)
                        if config.get("dir_abrir_pdf") and Path(config["dir_abrir_pdf"]).exists():
                            self.dir_abrir_pdf = config["dir_abrir_pdf"]
                        if config.get("dir_salvar_excel") and Path(config["dir_salvar_excel"]).exists():
                            self.dir_salvar_excel = config["dir_salvar_excel"]
        except Exception:
            pass  # Ignora erros de configuração — defaults já estão definidos no __init__

    def salvar_configuracoes(self, silencioso=False):
        """Salva configurações atuais."""
        try:
            config = {
                "dir_abrir_pdf": self.dir_abrir_pdf,
                "dir_salvar_excel": self.dir_salvar_excel,
                "data_ultima_execucao": datetime.now().isoformat()
            }
            config_file = Path(__file__).parent / "config.json"
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
    root = tk.Tk()
    app = PdfToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

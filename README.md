# Control ID Reader - Folha de Ponto para Excel

Converte um ou múltiplos arquivos de folha de ponto (PDF ou MHTML) da Control ID em uma única planilha Excel formatada.

## Funcionalidades

- ✅ Processa um ou múltiplos arquivos simultaneamente (PDF e MHTML)
- ✅ Suporte a PDFs com múltiplos funcionários em um único arquivo
- ✅ Separação de PDF multi-funcionário em arquivos individuais compactados em ZIP
- ✅ Detecção automática de múltiplos formatos de PDF (colunas separadas ou mescladas)
- ✅ Interface moderna com CustomTkinter — cantos arredondados, tipografia limpa
- ✅ Tema Dark/Light alternável com botão dedicado
- ✅ Botões com cores intuitivas (verde/laranja/vermelho/azul/roxo)
- ✅ Seleção de arquivos com preview visual
- ✅ Log de processamento em tempo real com cores por tipo de mensagem
- ✅ Barra de progresso visual
- ✅ Extrai dados completos do funcionário (nome, CPF, PIS, cargo, admissão, matrícula, centro de custo, departamento)
- ✅ Extrai marcações de ponto por dia (até 3 entradas/saídas), duração e carga horária
- ✅ Armazena horários contratuais do funcionário para uso futuro
- ✅ Gera planilha Excel formatada com cabeçalho, resumo e tabela de ponto ordenada
- ✅ Memoriza separadamente o último diretório usado para abrir arquivos e para salvar Excel
- ✅ Salva tema (dark/light) entre sessões

## Como usar

### Pré-requisitos

- Python 3.8 ou superior instalado (https://python.org)

### Passos

1. Abra o terminal nesta pasta
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Execute o script:
   ```bash
   python pdf_para_excel.py
   ```
   ou duplo-clique no arquivo no Explorer (Windows)

4. **Interface aberta:**
   - Clique em **"+ Adicionar Arquivos"** para selecionar PDFs ou MHTMLs (pode selecionar múltiplos)
   - Use **"✕ Remover Selecionados"** para remover apenas os arquivos marcados na lista
   - Use **"🗑 Limpar Tudo"** para limpar a lista inteira de uma vez
   - Clique em **"⚡ Gerar Planilha Excel"** para processar
   - Clique em **"✂️ Separar PDF por Funcionário"** para dividir um PDF com múltiplos funcionários

5. Escolha o local e nome para salvar a planilha ou o ZIP

6. Aguarde o processamento — o log mostrará o progresso em tempo real

## Separação de PDF por Funcionário

O botão **"✂️ Separar PDF por Funcionário"** permite dividir um único PDF que contenha múltiplos funcionários em arquivos individuais, compactados em um ZIP.

- Cada arquivo gerado recebe o nome do funcionário (sem acentos, espaços viram `_`)
- O processo é independente da lista de arquivos principal
- Após gerar o ZIP, os PDFs separados podem ser adicionados normalmente para conversão em Excel

> **Observação:** se o PDF contiver apenas um funcionário, o app avisa que a separação é desnecessária.

## Estrutura da planilha gerada

### Funcionário único (ou múltiplos arquivos do mesmo CPF)

A aba **Ponto** é organizada em três blocos:

1. **Cabeçalho** — dados do funcionário extraídos do arquivo (empresa, nome, CPF, PIS, matrícula, admissão, cargo, departamento, centro de custo). Campos não encontrados são omitidos automaticamente.

2. **Resumo** — linha única com dias efetivamente trabalhados (dias sem nenhuma marcação são excluídos) e média de horas por dia trabalhado.

3. **Tabela de ponto** — registros ordenados por data, com as colunas: DIA, MARCAÇÕES, ENT. 1, SAÍ. 1, ENT. 2, SAÍ. 2, ENT. 3, SAÍ. 3, DURAÇÃO, CH.

### Múltiplos funcionários

A aba **Ponto** contém a tabela completa com as colunas FUNCIONÁRIO e CPF adicionadas, ordenada por nome e depois por data. Uma aba separada **Resumo** é criada com uma linha por funcionário contendo: nome, CPF, PIS, cargo, admissão, matrícula, departamento, centro de custo, dias trabalhados e média de horas/dia.

### Nome do arquivo gerado

O nome é sugerido automaticamente com base no cenário:

| Cenário | Exemplo |
|---|---|
| 1 arquivo, 1 funcionário | `Ponto_Gabriela_Almeida_Abr2026.xlsx` |
| Vários arquivos, mesmo funcionário | `Ponto_Gabriela_Almeida.xlsx` |
| Vários funcionários | `Ponto_3Funcionarios_Mai2026.xlsx` |

## Formatos suportados

| Formato | Observação |
|---|---|
| `.pdf` | Folha de ponto exportada diretamente pelo sistema — suporta múltiplos layouts automaticamente |
| `.mhtml` | Salvo pelo navegador via "Salvar página como → Página da Web, Completa (*.mhtml)" |

> **Atenção:** arquivos `.html` simples geralmente não contêm dados (o conteúdo é carregado dinamicamente). Use sempre `.mhtml`.

## Configurações salvas automaticamente

O arquivo `config.json` é gerado automaticamente em `%APPDATA%\ControlIDReader\` e guarda:

```json
{
  "dir_abrir_pdf": "C:/Users/Fulano/Downloads",
  "dir_salvar_excel": "C:/Users/Fulano/Documents",
  "tema": "dark",
  "data_ultima_execucao": "2025-01-15T14:32:00"
}
```

> O caminho completo no Windows é `C:\Users\<seu_usuario>\AppData\Roaming\ControlIDReader\config.json`.

As configurações são salvas automaticamente sempre que você abre ou salva um arquivo. Na próxima execução, os diálogos já abrem nas últimas pastas usadas e o tema é restaurado.

O botão **"💾 Salvar Configurações"** permite forçar o salvamento manualmente com confirmação visual.

## Como criar o executável (.exe)

Opcionalmente, você pode gerar um executável para distribuir sem precisar do Python instalado:

### Usando o script automático (Windows)

```bash
./criar_exe.bat
```

### Manualmente

```bash
# Instale as dependências
pip install -r requirements.txt

# Gere o executável (Windows)
python -m PyInstaller --clean --onefile --windowed --name "Control ID Reader" ^
    --collect-all pdfplumber ^
    --collect-all pypdf ^
    --collect-all openpyxl ^
    --collect-all bs4 ^
    --collect-all customtkinter ^
    --icon "control_id_reader.ico" ^
    --add-data "control_id_reader.ico;." ^
    pdf_para_excel.py

# Ou use o script de atualização
./atualizar_exe.bat
```

Ou use o arquivo `.spec` incluído:

```bash
pyinstaller "Control ID Reader.spec"
```

O executável será criado em `dist/Control ID Reader.exe` (~55 MB).

## Estrutura do projeto

```
pdf_reader/
├── pdf_para_excel.py         # Script principal
├── requirements.txt          # Dependências do projeto
├── README.md                 # Este arquivo
├── criar_exe.bat             # Script para criar o .exe pela primeira vez
├── atualizar_exe.bat         # Script para recriar o .exe após alterações
├── control_id_reader.ico     # Ícone do app (janela e atalho)
├── Control ID Reader.spec    # Configuração do PyInstaller
└── dist/                     # Executável gerado (não versionado)
```

## Problemas conhecidos

### Windows bloqueia o executável

O Windows pode bloquear a execução do `.exe` por não ser assinado digitalmente. Para contornar:

1. Clique direito no arquivo `.exe` → **Propriedades**
2. Marque a caixa **"Desbloquear"** na parte inferior
3. Clique em **OK** e execute novamente

### Erro "ModuleNotFoundError"

Se o executável falhar com erro de módulo faltando, recrie o `.exe` usando o script `atualizar_exe.bat` — ele garante que todas as bibliotecas sejam incluídas corretamente.

### Erro de firewall do Windows

Se o Windows Defender Firewall bloquear o executável:
1. Abra **Configurações** → **Atualização e Segurança** → **Windows Security**
2. Clique em **Firewall e proteção de rede** → **Allow an app through firewall**
3. Clique em **Change settings** e adicione "Control ID Reader" à lista

## Licença

Uso livre.

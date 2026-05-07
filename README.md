# Control ID Reader - Folha de Ponto para Excel

Converte um ou múltiplos arquivos PDF de folha de ponto da Control ID em uma única planilha Excel formatada.

## Funcionalidades

- ✅ Processa um ou múltiplos PDFs simultaneamente
- ✅ Interface gráfica moderna, centralizada e responsiva
- ✅ Botões com cores intuitivas (verde/laranja/vermelho/azul)
- ✅ Seleção de arquivos com preview visual
- ✅ Log de processamento em tempo real
- ✅ Barra de progresso visual
- ✅ Extrai dados do funcionário (nome, CPF, PIS, cargo, admissão, matrícula, centro de custo, departamento)
- ✅ Extrai marcações de ponto por dia (3 entradas/saídas)
- ✅ Gera planilha Excel formatada com colunas ajustadas
- ✅ Memoriza separadamente o último diretório usado para abrir PDFs e para salvar Excel

## Como usar

### Pré-requisitos

- Python 3.8 ou superior instalado (https://python.org)
- **Observação:** `tkinter` já vem incluído com Python (não requer instalação separada)

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
   - Clique em **"+ Adicionar PDFs"** para selecionar arquivos (pode selecionar múltiplos)
   - Use **"❌ Remover Selecionados"** para remover apenas os arquivos marcados na lista
   - Use **"🗑️ Limpar Tudo"** para limpar a lista inteira de uma vez
   - Clique em **"Gerar Planilha Excel"** para processar

5. Escolha o local e nome para salvar a planilha

6. Aguarde o processamento — o log mostrará o progresso em tempo real

## Configurações salvas automaticamente

O arquivo `config.json` é gerado automaticamente na pasta do projeto e guarda:

```json
{
  "dir_abrir_pdf": "C:/Users/Fulano/Downloads",
  "dir_salvar_excel": "C:/Users/Fulano/Documents",
  "data_ultima_execucao": "2025-01-15T14:32:00"
}
```

Na próxima execução, os diálogos de abrir e salvar já abrem nas últimas pastas usadas.
O botão **"Salvar Configurações"** permite forçar o salvamento manualmente com confirmação visual.

## Como criar o executável (.exe)

Opcionalmente, você pode gerar um executável para distribuir:

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
    --collect-all openpyxl ^
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
├── Control ID Reader.spec    # Configuração do PyInstaller
├── config.json               # Configurações salvas (não versionado)
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

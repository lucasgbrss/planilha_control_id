# Control ID Reader

Converte um arquivo PDF de folha de ponto da Control ID em uma planilha Excel formatada.

## Funcionalidades

- Extrai dados de funcionários (nome, CPF, PIS, cargo, etc.)
- Extrai marcações de ponto por dia
- Gera planilha Excel formatada e pronta para uso
- Interface simples para seleção de arquivos

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
4. Selecione o PDF de ponto desejado
5. Selecione o local onde a planilha será salva

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

# Gere o executável
python -m PyInstaller --clean --onefile --windowed --name "Control ID Reader" ^
    --collect-all pdfplumber ^
    --collect-all openpyxl ^
    --collect-all pypdfium2 ^
    --collect-all Pillow ^
    pdf_para_excel.py
```

Ou use o arquivo `.spec` incluído:

```bash
pyinstaller Control ID Reader.spec
```

O executável será criado em `dist/Control ID Reader.exe` (~55 MB).

## Estrutura do projeto

```
pdf_reader/
├── pdf_para_excel.py      # Script principal
├── requirements.txt       # Dependências do projeto
├── README.md              # Este arquivo
├── criar_exe.bat          # Script para criar o .exe
├── Control ID Reader.spec # Configuração do PyInstaller
└── dist/                  # Executável gerado (não versionado)
```

## Problemas conhecidos

### Windows bloqueia o executável

O Windows pode bloquear a execução do `.exe` por não ser assinado digitalmente. Para contornar:

1. Clique direito no arquivo `.exe` → **Propriedades**
2. Marque a caixa **"Desbloquear"** na parte inferior
3. Clique em **OK** e execute novamente

### Erro "ModuleNotFoundError"

Se o executável falhar com erro de módulo faltando, recrie o `.exe` usando o script `criar_exe.bat` — ele garante que todas as bibliotecas sejam incluídas corretamente.

## Licença

Uso livre.

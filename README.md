# PDF de Ponto para Excel

Converte um arquivo PDF de folha de ponto da Control ID em uma planilha Excel formatada.

## Como usar

### Opção 1: Executável (.exe) - Recomendado para usuários finais

1. Baixe o arquivo `PDF_Ponto_Excel.exe` da pasta `dist/`
2. Execute o arquivo
3. Selecione o PDF de ponto desejado
4. A planilha será gerada na pasta **Documentos**

> Não requer Python instalado!

### Opção 2: Rodar o script Python

**Pré-requisitos:** Python 3.8 ou superior instalado (https://python.org)

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
5. A planilha será gerada na pasta **Documentos**

## Como criar o executável (.exe)

Se você quer gerar o executável para distribuir:

```bash
# Instale as dependências
pip install -r requirements.txt

# Gere o executável (usa o comando correto para Windows)
python -m PyInstaller --clean --onefile --windowed --name "PDF_Ponto_Excel" ^
    --collect-all pdfplumber ^
    --collect-all openpyxl ^
    --collect-all pypdfium2 ^
    --collect-all Pillow ^
    pdf_para_excel.py
```

Ou use o script automático `criar_exe.bat` (Windows).

O executável será criado em `dist/PDF_Ponto_Excel.exe` (~55 MB).

## Estrutura do projeto

```
pdf_reader/
├── pdf_para_excel.py      # Script principal
├── requirements.txt       # Dependências do projeto
├── README.md              # Este arquivo
├── criar_exe.bat          # Script para criar o .exe
└── dist/
    └── PDF_Ponto_Excel.exe  # Executável gerado
```

## Problemas conhecidos

### Windows bloqueia o executável

O Windows pode bloquear a execução do `.exe` por não ser assinado digitalmente. Para contornar:

1. Clique direito no arquivo `.exe` → **Propriedades**
2. Marque a caixa **"Desbloquear"** na parte inferior
3. Clique em **OK** e execute novamente

### Erro "ModuleNotFoundError"

Se o executável falhar com erro de módulo faltando, recrie o `.exe` usando o script `criar_exe.bat` — ele garante que todas as bibliotecas sejam incluídas corretamente.

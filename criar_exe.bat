@echo off
echo ========================================
echo  Criando executavel PDF Ponto para Excel
echo ========================================
echo.

echo Instalando dependencias...
pip install -r requirements.txt

echo.
echo Gerando executavel com PyInstaller...
REM Usa python -m para garantir que usa o Python correto
python -m PyInstaller --clean --onefile --windowed --name "Control ID Reader" --collect-all pdfplumber --collect-all openpyxl --collect-all pypdfium2 --collect-all Pillow pdf_para_excel.py

echo.
echo ========================================
echo  Concluido!
echo  O executavel esta em: dist\PDF_Ponto_Excel.exe
echo ========================================
pause

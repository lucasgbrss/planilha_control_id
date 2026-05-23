@echo off
echo ========================================
echo  Criando executavel PDF Ponto para Excel
echo ========================================
echo.

echo Instalando dependencias...
pip install -r requirements.txt

echo.
echo Gerando executavel com PyInstaller...
python -m PyInstaller --clean --onefile --windowed --name "Control ID Reader" ^
    --collect-all pdfplumber ^
    --collect-all openpyxl ^
    --collect-all bs4 ^
    --icon "control_id_reader.ico" ^
    --add-data "control_id_reader.ico;." ^
    pdf_para_excel.py

echo.
echo ========================================
echo  Concluido!
echo  O executavel esta em: dist\Control ID Reader.exe
echo ========================================
pause

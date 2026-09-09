@echo off
echo ========================================
echo  Atualizando executavel PDF Ponto para Excel
echo ========================================
echo.

echo Gerando executavel com PyInstaller...
python -m PyInstaller --clean --onefile --windowed --name "Control ID Reader" ^
    --collect-all pdfplumber ^
    --collect-all pypdf ^
    --collect-all openpyxl ^
    --collect-all bs4 ^
    --collect-all customtkinter ^
    --collect-all cryptography ^
    --icon "control_id_reader.ico" ^
    --add-data "control_id_reader.ico;." ^
    --distpath dist ^
    --noconfirm ^
    pdf_para_excel.py

echo.
echo ========================================
echo  Concluido!
echo  O executavel esta em: dist\Control ID Reader.exe
echo ========================================
pause

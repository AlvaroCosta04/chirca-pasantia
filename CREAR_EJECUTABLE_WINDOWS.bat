@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ===========================================================================
echo   EMPAQUETADOR AUTOMATICO (.EXE) - SISTEMA DE CALIDAD CHIRCA OVC S.A.
echo ===========================================================================

echo.
echo [1/4] Instalando librerias necesarias y PyInstaller...
python -m pip install --upgrade pip
pip install pandas numpy pymupdf rapidocr-onnxruntime python-docx openpyxl matplotlib pyinstaller

echo.
echo [2/4] Compilando ACTUALIZAR_INFORMES.exe (esto demora unos 2 a 3 minutos)...
pyinstaller --noconfirm --onedir --console --name "ACTUALIZAR_INFORMES" ^
    --collect-all rapidocr_onnxruntime ^
    --collect-all pymupdf ^
    --collect-all docx ^
    --collect-all openpyxl ^
    --collect-all matplotlib ^
    EJECUTAR_SISTEMA.py

echo.
echo [3/4] Copiando bases de datos CSV y scripts dentro de la carpeta portable...
xcopy /E /I /Y "01_Bases_de_Datos_CSV" "dist\ACTUALIZAR_INFORMES\01_Bases_de_Datos_CSV" > nul
xcopy /E /I /Y "03_Sistema_y_Scripts" "dist\ACTUALIZAR_INFORMES\03_Sistema_y_Scripts" > nul

echo.
echo [4/4] Limpiando archivos temporales de compilacion...
rmdir /S /Q build 2>nul
del /Q ACTUALIZAR_INFORMES.spec 2>nul

echo.
echo ===========================================================================
echo   ¡LISTO! La carpeta portable con el .EXE quedo creada en:
echo   dist\ACTUALIZAR_INFORMES\ACTUALIZAR_INFORMES.exe
echo ===========================================================================
pause

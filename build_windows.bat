@echo off
setlocal

rem Build locale dell'eseguibile onedir. L'installer NSIS viene creato separatamente.
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean --windowed ^
  --name "CollaudoSuite" ^
  --icon "collaudo_suite\assets\app_icon.ico" ^
  --collect-all matplotlib ^
  --collect-all nltk ^
  --hidden-import PySide6.QtPrintSupport ^
  --hidden-import PySide6.QtPdf ^
  --hidden-import PySide6.QtPdfWidgets ^
  --hidden-import collaudo_suite.checklist.data ^
  --add-data "collaudo_suite\checklist\data;collaudo_suite\checklist\data" ^
  --add-data "collaudo_suite\assets\Guida_operativa_Collaudo_Suite.pdf;collaudo_suite\assets" ^
  --add-data "collaudo_suite\assets\app_icon.png;collaudo_suite\assets" ^
  --add-data "version.txt;." ^
  run_suite.py

if errorlevel 1 exit /b %errorlevel%
echo Build completata in dist\CollaudoSuite
endlocal

@echo off
setlocal
cd /d "%~dp0"
echo Budowanie aplikacji Procesy na sztuke...
where py >nul 2>nul
if errorlevel 1 (
  echo Zainstaluj Python 3.12 z python.org i uruchom ten plik ponownie.
  pause
  exit /b 1
)
if not exist .venv-build\Scripts\python.exe (
  py -3.12 -m venv .venv-build
  if errorlevel 1 goto failed
)
.venv-build\Scripts\python.exe -m pip install -r requirements.txt pyinstaller==6.16.0
if errorlevel 1 goto failed
.venv-build\Scripts\python.exe -m unittest discover -s tests -v
if errorlevel 1 goto failed
.venv-build\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name Procesy-na-sztuke --add-data "static:static" desktop.py
if errorlevel 1 goto failed
echo Gotowe! Plik: dist\Procesy-na-sztuke.exe
explorer "%cd%\dist"
pause
exit /b 0
:failed
echo Budowanie nie powiodlo sie. Sprawdz komunikat powyzej.
pause
exit /b 1

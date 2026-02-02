@echo off
cd /d "%~dp0"
echo Instalando dependencias...
py -3 -m pip install -r requirements.txt -q
py -3 -m pip install pyinstaller -q
echo Generando ejecutable...
py -3 -m PyInstaller GeneradorEssar.spec --noconfirm --distpath . --workpath build
echo Listo. Ejecutable: %~dp0Generador Essar.exe
pause

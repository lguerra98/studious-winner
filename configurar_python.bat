@echo off

if not exist "%cd%\.venv\" (
    py -m venv .venv
)

call "%cd%\.venv\Scripts\activate.bat"
python -m pip install -r .\requirements.txt
echo Configuracion creada exitosamente

pause
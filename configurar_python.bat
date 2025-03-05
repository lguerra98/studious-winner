@echo off

if not exists "%cd%\.venv\"(
    python -m venv .venv
)

call "%cd%\.venv\Scripts\activate.bat"
python -m pip install -r .\requirements.txt
echo Configuracion creada exitosamente
pause
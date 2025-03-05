@echo off

call "%cd%\.venv\Scripts\activate.bat"
python "%cd%\run2.py"
echo Factura generada existosamente
pause
@echo off

cd /d "C:\Users\%USERNAME%\Desktop\New Projects\Medico\Medico"
python -m virtualenv venv

cd venv/scripts
call activate.bat

cd /d "C:\Users\%USERNAME%\Desktop\New Projects\Medico\Medico"
pip install -r requirements.txt

cmd /k
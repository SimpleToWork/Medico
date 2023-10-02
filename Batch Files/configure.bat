@echo off

C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\scripts\pip install virtualenv

cd /d "C:\Users\%USERNAME%\Desktop\New Projects\Medico\Medico"
C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\python -m virtualenv venv

cd venv/scripts
call activate.bat

cd /d "C:\Users\%USERNAME%\Desktop\New Projects\Medico\Medico"
pip install -r requirements.txt

cmd /k
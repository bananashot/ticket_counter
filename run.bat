@echo off

set current_directory=%cd%

cmd /k "cd /d %current_directory%\venv\Scripts & activate & cd /d    %current_directory% & python main.py"
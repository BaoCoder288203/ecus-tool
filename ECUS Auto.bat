@echo off
cd /d "%~dp0"
where pythonw >NUL 2>NUL && (start "" pythonw ecus_auto.py) || (python ecus_auto.py & pause)

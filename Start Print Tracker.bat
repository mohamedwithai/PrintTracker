@echo off
title Print Tracker
cd /d "%~dp0"
set "CONDA=%USERPROFILE%\anaconda3"
if not exist "%CONDA%\python.exe" goto syspy
set "PATH=%CONDA%;%CONDA%\Library\bin;%CONDA%\Scripts;%PATH%"
"%CONDA%\python.exe" app.py
goto done
:syspy
where python >nul 2>nul
if errorlevel 1 goto nopy
python app.py
goto done
:nopy
echo Python was not found on this PC. Install Python from python.org and try again.
:done
echo.
echo Print Tracker has stopped.
pause

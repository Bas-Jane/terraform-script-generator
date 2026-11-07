@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%~dp0ibr_tfvars_generator.py" %*
) else (
  python "%~dp0ibr_tfvars_generator.py" %*
)
if errorlevel 1 pause

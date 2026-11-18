@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 scripts\build_dashboard.py
) else (
  python scripts\build_dashboard.py
)
if errorlevel 1 (
  echo.
  echo Не удалось пересоздать дашборд. Проверьте, что установлен Python 3.
  pause
  exit /b 1
)
echo.
echo Готово. Теперь откройте index.html.
pause

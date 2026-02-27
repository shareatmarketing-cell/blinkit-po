@echo off
echo Starting PO Blinkit Extractor...

:: Start backend
start "PO Backend" cmd /k "cd /d "%~dp0backend" && uvicorn main:app --reload --port 8000"

:: Give backend a moment to start
timeout /t 2 /nobreak > nul

:: Start frontend
start "PO Frontend" cmd /k "cd /d "%~dp0frontend" && npm run dev"

echo.
echo Backend:  http://localhost:8000
echo Frontend: http://localhost:5173
echo.
pause

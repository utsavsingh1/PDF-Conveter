@echo off
echo ============================================
echo   PDF to Any - PaddleOCR Converter
echo ============================================
echo.
echo  Starting Flask on http://127.0.0.1:5000 ...
echo  DO NOT open index.html directly in browser!
echo  Always use: http://127.0.0.1:5000
echo.
start "PDFAny Backend" cmd /k "cd /d "%~dp0" && python app.py"
echo  Waiting for server to start...
timeout /t 4 /nobreak >nul
start http://127.0.0.1:5000
echo.
echo  Server running. Press any key to stop.
pause >nul
taskkill /FI "WindowTitle eq PDFAny Backend*" /T /F >nul 2>&1
echo  Server stopped.

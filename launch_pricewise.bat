@echo off
echo ==========================================
echo    PriceWise - Universal Startup Script
echo ==========================================

:: 1. Start Python Processing Microservice (Minimized)
echo [1/3] Launching Python Matcher Service...
start /min "PriceWise Matcher" cmd /k "cd /d "%~dp0services\processing" && python main.py"

:: 2. Start Node.js API Gateway (Minimized)
echo [2/3] Launching API Gateway...
start /min "PriceWise API" cmd /k "cd /d "%~dp0apps\api" && npm run dev"

:: 3. Start Next.js Frontend (Minimized)
echo [3/3] Launching Web Frontend...
start /min "PriceWise Web" cmd /k "cd /d "%~dp0web" && npm run dev"

:: 4. Launch Browser (wait longer so dev servers have time to boot)
echo [4/4] Opening PriceWise in your browser...
timeout /t 12 >nul
start http://localhost:3000

echo ==========================================
echo    All services are running in background!
echo    Web UI:  http://localhost:3000
echo    API:     http://localhost:3001
echo ==========================================
timeout /t 3 >nul
exit

@echo off
TITLE Cat Test Case Builder - Starting...
echo --------------------------------------------------
echo   Cat Test Case Builder ฅ^•ﻌ•^ฅ
echo   Starting the server...
echo --------------------------------------------------

REM Check if node_modules exists, if not run npm install
if not exist "node_modules\" (
    echo [1/3] node_modules not found. Installing dependencies...
    call npm install
) else (
    echo [1/3] Dependencies already installed.
)

echo [2/3] Starting server...
REM Start the server in a new window so this script can continue to open the browser
start /b node index.js

echo [3/3] Opening browser at http://localhost:3000...
timeout /t 3 >nul
start http://localhost:3000

echo --------------------------------------------------
echo   App is running! 🐾
echo   Don't close this window while using the app.
echo --------------------------------------------------
pause

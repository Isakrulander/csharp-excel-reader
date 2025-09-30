@echo off
echo Starting Excel Data Analysis Platform...
echo.

echo [1/2] Starting C# Backend API...
cd /d "C:\Users\isakr\OneDrive - BTH Student\Codenr3\NKTCS\WebAPI"
start "Backend API" cmd /k "dotnet run"

echo [2/2] Starting Vue.js Frontend...
cd /d "C:\Users\isakr\OneDrive - BTH Student\Codenr3\NKTCS\Frontend"
start "Frontend" cmd /k "npm run dev"

echo.
echo ==========================================
echo   Excel Data Analysis Platform Started
echo   Backend:  http://localhost:5000
echo   Frontend: http://localhost:3000
echo ==========================================
echo.
echo Press any key to continue...
pause > nul
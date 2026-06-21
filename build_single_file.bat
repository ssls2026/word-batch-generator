@echo off
echo ============================================================
echo   Word Batch Generator - Single File EXE Build Script
echo ============================================================
echo.

echo [1/2] Cleaning old build files...
if exist "%~dp0dist\Word批量生成器.exe" (
    del /f /q "%~dp0dist\Word批量生成器.exe"
)

echo [2/2] Publishing standalone single-file EXE (Release)...
cd /d "%~dp0WordBatchGenerator"
dotnet publish -c Release --nologo
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed! Please ensure .NET 8 SDK is installed.
    pause
    exit /b 1
)
cd /d "%~dp0"

echo.
echo ============================================================
echo   Single File EXE generated successfully: dist\Word批量生成器.exe
echo ============================================================
echo.
explorer "%~dp0dist"
pause

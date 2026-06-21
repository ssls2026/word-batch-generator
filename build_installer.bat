@echo off
chcp 65001 >nul
echo ============================================================
echo   Word 批量生成器 - 安装包构建脚本
echo ============================================================
echo.

set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
set ISS_FILE=%~dp0installer\setup.iss
set DIST_EXE=%~dp0dist\Word批量生成器.exe

:: 1. 检查 Inno Setup
if not exist %ISCC% (
    echo [错误] 未找到 Inno Setup 6，请先安装：
    echo   https://jrsoftware.org/isdl.php
    pause & exit /b 1
)

:: 2. 检查源 exe 是否存在
if not exist "%DIST_EXE%" (
    echo [1/2] 源 exe 不存在，先执行 dotnet publish...
    cd /d "%~dp0WordBatchGenerator"
    dotnet publish -c Release --nologo
    if errorlevel 1 ( echo [错误] publish 失败 & pause & exit /b 1 )
    cd /d "%~dp0"
) else (
    echo [1/2] 使用已有 dist\Word批量生成器.exe
)

:: 3. 编译安装包
echo [2/2] 正在编译安装包...
%ISCC% "%ISS_FILE%"
if errorlevel 1 ( echo [错误] 安装包编译失败 & pause & exit /b 1 )

echo.
echo ============================================================
echo   安装包已生成：installer\output\Word批量生成器_v2.0_Setup.exe
echo ============================================================
echo.
explorer "%~dp0installer\output"
pause

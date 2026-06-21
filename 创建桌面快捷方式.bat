@echo off
chcp 65001 >nul
:: 创建桌面快捷方式（解决大型单文件exe在桌面图标显示问题）

set EXE_PATH=%~dp0dist\Word批量生成器.exe
set SHORTCUT_NAME=Word批量生成器
set DESKTOP=%USERPROFILE%\Desktop

if not exist "%EXE_PATH%" (
    echo [错误] 未找到: %EXE_PATH%
    echo 请先运行 dotnet publish -c Release 生成 dist\Word批量生成器.exe
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$sc = $ws.CreateShortcut('%DESKTOP%\%SHORTCUT_NAME%.lnk');" ^
  "$sc.TargetPath = '%EXE_PATH%';" ^
  "$sc.IconLocation = '%EXE_PATH%, 0';" ^
  "$sc.WorkingDirectory = '%~dp0dist';" ^
  "$sc.Description = 'Word 批量生成器 v2.0';" ^
  "$sc.Save();" ^
  "Write-Host '快捷方式已创建到桌面'"

echo.
echo [完成] 桌面快捷方式已创建：%SHORTCUT_NAME%.lnk
pause

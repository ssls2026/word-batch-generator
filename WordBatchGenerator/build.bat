@echo off
chcp 65001 >nul
echo ========================================
echo   Word 批量生成器 - 一键构建脚本
echo ========================================
echo.

echo [1/3] 清理旧版本...
if exist "bin\Release" rd /s /q "bin\Release"
if exist "obj\Release" rd /s /q "obj\Release"

echo [2/3] 还原依赖包...
dotnet restore

echo [3/3] 发布单文件版本...
dotnet publish -c Release -r win-x64 --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -p:EnableCompressionInSingleFile=true

echo.
echo ========================================
echo   ✅ 构建完成！
echo ========================================
echo.
echo 输出文件位置:
echo bin\Release\net8.0\win-x64\publish\Word批量生成器.exe
echo.
echo 文件大小: 约 60-80 MB (WPF 不支持裁剪)
echo 启动速度: < 1 秒
echo 依赖要求: 无（可直接分发）
echo.
pause

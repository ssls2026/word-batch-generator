@echo off
chcp 65001 >nul
echo ============================================================
echo          Word 批量生成器 - 单文件打包脚本 (WPF .NET 8)
echo ============================================================
echo.

echo [1/3] 清理历史 Release 构建缓存...
if exist "WordBatchGenerator\bin\Release" rd /s /q "WordBatchGenerator\bin\Release"
if exist "WordBatchGenerator\obj\Release" rd /s /q "WordBatchGenerator\obj\Release"

echo [2/3] 还原项目 NuGet 依赖项...
dotnet restore WordBatchGenerator\WordBatchGenerator.csproj

echo [3/3] 执行单文件发布编译 (Self-Contained + 压缩)...
dotnet publish WordBatchGenerator\WordBatchGenerator.csproj -c Release -r win-x64 --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -p:EnableCompressionInSingleFile=true

echo.
echo [4/4] 正在将单文件 EXE 复制到根目录...
copy "WordBatchGenerator\bin\Release\net8.0-windows\win-x64\publish\Word批量生成器.exe" "Word批量生成器.exe" /y

echo.
echo ============================================================
echo   ✅ 恭喜！单文件编译构建已完成！
echo ============================================================
echo.
echo 根目录单文件可执行程序路径:
echo Word批量生成器.exe
echo.
echo 构建信息说明:
echo * 文件大小: 约 72 MB 左右 (已集成 .NET 8 离线环境与所有第三方依赖库)
echo * 运行平台: Windows x64 系统 (Win10/Win11 绿色免安装运行)
echo * 启动机制: 双击后自动将原生 C++ 核心库自解压到临时目录并极速启动
echo.
pause

# Word 批量生成器 - 单文件打包与发布指南

本项目基于 **.NET 8.0 WPF** 进行重构开发。为了实现“无环境依赖”、“绿色免安装”、“单文件分发”的目标，我们使用 .NET 8 的高级发布（Publish）技术进行打包。

---

## 🚀 快速一键构建

我们在项目根目录和子目录下都提供了一键打包脚本：
* **根目录下**：[build_single_file.bat](file:///e:/Code/word-批量工具-csharp重构/build_single_file.bat) （推荐使用）
* **项目目录下**：[WordBatchGenerator/build.bat](file:///e:/Code/word-批量工具-csharp重构/WordBatchGenerator/build.bat)

### 使用步骤：
1. 双击运行根目录下的 `build_single_file.bat`。
2. 脚本会自动清理旧的构建缓存、还原 NuGet 依赖并执行发布。
3. 打包完成后，您可以在以下目录中找到打包好的单文件 exe：
   `WordBatchGenerator/bin/Release/net8.0-windows/win-x64/publish/Word批量生成器.exe`
4. 将此 `.exe` 发送给任意 Windows 64位系统用户即可，用户电脑上**不需要安装任何 .NET 运行时**。

---

## 🛠️ 底层构建命令说明

如果您想在终端手动构建，请在命令行中执行以下命令（在 `WordBatchGenerator` 目录下或在根目录下指定项目文件）：

```bash
dotnet publish WordBatchGenerator/WordBatchGenerator.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```

### 关键编译参数解析：

| 参数项 | 参数值/开关 | 作用与原理说明 |
| :--- | :--- | :--- |
| `-c` | `Release` | 采用 **Release** 模式进行优化编译，移除调试符号并对代码运行效率进行深度优化。 |
| `-r` | `win-x64` | 指定目标运行时为 **Windows 64位** 架构平台。 |
| `--self-contained` | `true` | **独立运行模式**。把整个 .NET 8.0 运行时核心打包进可执行文件中，因此目标机器无需安装任何 .NET 环境即可开箱即用。 |
| `-p:PublishSingleFile` | `true` | **单文件打包**。将所有的托管程序集（C# dll）、配置文件等打包合并为一个单一的 `.exe` 格式文件。 |
| `-p:IncludeNativeLibrariesForSelfExtract` | `true` | **本地原生库自解压提取**。将 WebView2 底层核心原生 DLL (`WebView2Loader.dll`) 等原生二进制文件一同嵌入单文件中。运行时它们会自动解压到用户的临时目录以被正常调用。 |
| `-p:EnableCompressionInSingleFile` | `true` | **启用单文件压缩**。通过内置的压缩算法，对合并后的单文件体积进行高比例压缩（体积从原本的 ~150MB 骤降至约 **72MB** 左右），极大地方便了网络传输和分发。 |

---

## 📝 绿色运行与持久化规范

本重构版本做到了真正的绿色运行。无论您将其放置在只读目录下还是只读 U 盘中，程序都会保证完美高可用：
1. **数据与方案目录**：所有创建的方案（包括配置及 Word 模板）默认存储在当前登录用户的 `AppData/Local/WordBatchGenerator/Schemes` 中。
2. **WebView2 缓存文件夹**：WebView2 渲染缓存自动指向 `AppData/Local/WordBatchGenerator/WebView2_Cache`，100% 避免由于目录只读或沙盒导致预览界面崩溃的隐患。
3. **全局配置与状态**：用户的窗口大小、位置、上次活动界面、上一次使用的方案信息自动保存在 `AppData/Local/WordBatchGenerator/Schemes/last_state.json` 中，软件启动时自动无缝还原。

# Word 批量生成器（C# 重构版）

基于 .NET 8 WPF 构建的 Word 批量变量替换与文件生成工具，支持单文件绿色部署。

---

## 编译为单文件 exe

**环境要求**：安装 [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)（运行时无需安装）

### 一键发布

```bash
cd WordBatchGenerator
dotnet publish -c Release
```

发布完成后，exe 输出到项目根目录下的 **`dist\Word批量生成器.exe`**（约 74 MB，自包含，无需任何运行时）。

### 更换图标

如需重新生成 ICO 文件（修改了 `WordBatchGenerator/Resources/app.png` 后）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\make_ico.ps1
```

---

## 技术栈

| 组件 | 说明 |
|------|------|
| .NET 8 + WPF | 主框架 + UI |
| ModernWpfUI | 现代 UI 控件库 |
| DocumentFormat.OpenXml | Word 文档处理 |
| EPPlus | Excel 数据读取 |
| Microsoft WebView2 | 模板 HTML 预览 |

---

## 项目结构

```
WordBatchGenerator/
├── Core/
│   ├── WordParser.cs        # Word → HTML 解析与预览
│   ├── ExcelHandler.cs      # Excel 数据读取
│   ├── Generator.cs         # 批量生成引擎
│   └── SchemeManager.cs     # 方案管理（持久化）
├── Gui/
│   ├── Panels/
│   │   ├── NewSchemePanel.xaml      # 新建/编辑方案
│   │   ├── SchemeListPanel.xaml     # 方案列表
│   │   └── GeneratePanel.xaml       # 批量生成
│   └── Themes/                      # 样式资源
├── Resources/
│   ├── app.ico              # 应用图标（多分辨率 ICO）
│   └── app.png              # 图标源图（用于重新生成 ICO）
├── MainWindow.xaml          # 主窗口（侧边栏导航）
└── App.xaml                 # 应用入口 + 全局资源

scripts/
└── make_ico.ps1             # PNG → 多分辨率 ICO 转换工具

dist/
└── Word批量生成器.exe        # 发布输出（不纳入 Git）
```

---

## 数据存储

程序数据全部存储在用户目录，不写入 exe 所在目录，支持放在只读位置运行：

```
%LOCALAPPDATA%\WordBatchGenerator\
├── Schemes\              # 方案数据（config.json + template.docx）
│   └── last_state.json   # 上次窗口状态 / 活动方案（自动恢复）
└── WebView2_Cache\       # WebView2 渲染缓存
```

---

## 功能说明

1. **方案管理** — 创建 / 编辑 / 删除方案，导入导出 `.wsp` 方案包
2. **模板配置** — 导入 Word 模板，可视化预览，标注 `{{变量名}}` 占位符
3. **批量生成** — 从 Excel 读取数据行，逐行替换变量，输出到指定目录

---

*LZS · 2026 · MIT License*

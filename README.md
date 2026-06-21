# Word 批量生成器 v2.0

> 基于 .NET 8 + WPF 构建的 Word 批量变量替换与文件生成工具。从 Excel 数据源批量生成 Word 文档，单文件绿色部署，无需安装任何运行时。

---

## 功能特性

### 🗂️ 方案管理
- 创建、编辑、删除方案，每个方案独立存储配置与模板
- **导入 / 导出 `.wsp` 方案包**：可将方案打包分享给他人，或跨机器迁移
- 自动记录上次使用的方案，下次启动直接还原

### 📄 模板配置
- 导入 Word 模板（`.docx`），在预览区实时查看排版效果
- 使用 `{{变量名}}` 语法标注占位符，支持在正文、页眉、页脚中使用
- 变量列表自动从模板中解析，与 Excel 表头一一对应

### ⚡ 批量生成
- 从 Excel 读取数据（`.xlsx`），每行生成一份 Word 文件
- 自定义文件名模板，支持变量插值（如 `授权书-{{公司名称}}.docx`）
- 支持按变量字段自动创建子文件夹分类输出
- 自动注入 `{{序号}}` 变量（无需 Excel 中预先设置）
- 实时进度显示，生成完成后自动打开输出目录

### 🖨️ 预览与打印
- Word 模板 HTML 预览（基于 WebView2）
- 生成后支持预览、打印单份文档

### 💾 数据持久化
- 所有配置写入用户 `AppData`，支持放在只读目录或 U 盘运行
- 窗口位置、大小、当前面板自动保存，下次启动无缝恢复

---

## 快速开始（用户）

### 方式一：安装版（推荐，图标显示最稳定）

直接运行 `installer\output\Word批量生成器_v2.0_Setup.exe`，按向导安装即可。安装后自动创建桌面快捷方式。

### 方式二：绿色单文件

将 `dist\Word批量生成器.exe` 复制到任意目录，双击运行，无需安装任何依赖。

> **注意**：若将 exe 直接放在桌面，Windows 对大文件（>50MB）的桌面图标提取有超时机制，可能显示默认图标。建议使用安装版，或运行 `创建桌面快捷方式.bat` 在桌面创建快捷方式。

---

## 使用说明

### 第一步：创建方案

1. 点击左侧「**新建方案**」
2. 填写方案名称
3. 点击「**导入 Word 模板**」，选择 `.docx` 文件
4. 在预览区确认模板内容，检查 `{{变量名}}` 占位符是否正确
5. 点击「**保存方案**」

### 第二步：批量生成

1. 点击左侧「**批量生成**」
2. 选择要使用的方案
3. 点击「**选择 Excel 数据源**」，导入 `.xlsx` 文件（表头需与模板变量名一致）
4. 设置输出目录
5. 设置文件名模板（例如：`{{公司名称}}_授权书.docx`）
6. 点击「**开始批量生成**」，等待完成

### Excel 数据格式

| 公司名称 | 联系人 | 有效期 |
|---------|--------|--------|
| XX 科技有限公司 | 张三 | 2026-12-31 |
| YY 贸易有限公司 | 李四 | 2026-06-30 |

- 第一行为表头，对应模板中的 `{{变量名}}`
- 每一行数据生成一份文件
- 无需手动添加序号列，系统自动注入 `{{序号}}`

---

## 开发者指南

### 环境要求

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- Windows 10 / 11（WPF 仅支持 Windows）
- Visual Studio 2022 或 VS Code（可选）

### 技术栈

| 组件 | 版本 | 用途 |
|------|------|------|
| .NET 8 + WPF | 8.0 LTS | 主框架 + UI |
| ModernWpfUI | 0.9.6 | 现代 Fluent 风格控件 |
| DocumentFormat.OpenXml | 3.0.2 | Word 文档读写 |
| EPPlus | 7.1.3 | Excel 数据读取 |
| Microsoft WebView2 | 1.0.2535 | 模板 HTML 预览 |

### 项目结构

```
word-批量工具-csharp重构/
│
├── WordBatchGenerator/              # 主项目
│   ├── Core/
│   │   ├── WordParser.cs            # Word → HTML 解析、变量提取、预览渲染
│   │   ├── ExcelHandler.cs          # Excel 读取、表头解析、数据验证
│   │   ├── Generator.cs             # 批量生成引擎（变量替换、子文件夹、进度回调）
│   │   └── SchemeManager.cs         # 方案 CRUD、.wsp 包导入导出、状态持久化
│   ├── Gui/
│   │   ├── Panels/
│   │   │   ├── NewSchemePanel.xaml  # 新建/编辑方案、模板预览、变量标注
│   │   │   ├── SchemeListPanel.xaml # 方案列表、导入导出操作
│   │   │   └── GeneratePanel.xaml   # 批量生成、进度显示、结果日志
│   │   ├── ModernMessageBox.xaml    # 自定义消息对话框
│   │   ├── PrintWindow.xaml         # 打印预览窗口
│   │   └── Themes/                  # 全局样式（颜色、字体、控件模板）
│   ├── Resources/
│   │   ├── app.ico                  # 应用图标（BMP 16/32/48 + PNG 256，多分辨率）
│   │   └── app.png                  # 图标源图（蓝色主题，用于重新生成 ICO）
│   ├── MainWindow.xaml              # 主窗口（深色侧边栏导航）
│   └── WordBatchGenerator.csproj   # 项目配置（单文件发布、图标、依赖）
│
├── installer/
│   └── setup.iss                    # Inno Setup 安装包脚本
│
├── scripts/
│   ├── make_ico.ps1                 # PNG → 多分辨率 ICO 转换脚本
│   └── check_icon.cs               # 调试用：检查 PE 文件图标资源
│
├── dist/                            # 发布输出（gitignore，不入库）
│   └── Word批量生成器.exe
│
├── build_installer.bat              # 一键构建安装包脚本
└── 创建桌面快捷方式.bat              # 绿色版用户创建桌面快捷方式
```

### 编译与发布

#### 开发模式运行

```bash
cd WordBatchGenerator
dotnet run
```

#### 发布单文件 exe

```bash
cd WordBatchGenerator
dotnet publish -c Release
# 输出：../dist/Word批量生成器.exe（约 74 MB，自包含）
```

#### 构建安装包

```bat
# 双击运行（需已安装 Inno Setup 6）
build_installer.bat
# 输出：installer/output/Word批量生成器_v2.0_Setup.exe（约 68 MB）
```

或手动编译：
```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
```

#### 更换图标

修改 `WordBatchGenerator/Resources/app.png` 后，运行以下命令重新生成 ICO（BMP 多分辨率格式，Inno Setup 和 Windows Shell 兼容）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\make_ico.ps1
```

然后重新 publish 使新图标嵌入 exe。

---

## 数据存储

程序所有数据写入用户目录，不写入 exe 所在位置，支持只读目录和 U 盘运行：

```
%LOCALAPPDATA%\WordBatchGenerator\
├── Schemes\
│   ├── <方案名>\
│   │   ├── config.json      # 方案配置（变量列表、输出设置等）
│   │   └── template.docx    # Word 模板副本
│   └── last_state.json      # 上次窗口状态、活动方案（启动时自动恢复）
└── WebView2_Cache\          # WebView2 渲染缓存（卸载时可自动清理）
```

---

## 常见问题

**Q：双击 exe 没反应 / 预览区空白**
> 需要安装 WebView2 运行时。程序会自动检测并弹出下载提示。大多数 Windows 10（20H2+）和 Windows 11 已内置，无需手动安装。

**Q：桌面图标显示为默认图标**
> Windows Shell 对超过 50MB 的 exe 提取图标时可能超时。解决方法：
> 1. 推荐使用**安装版**（`build_installer.bat` 构建）
> 2. 或运行 `创建桌面快捷方式.bat` 创建快捷方式放桌面

**Q：Excel 数据读取报错**
> 请确保 Excel 文件未被 WPS/Office 打开，且文件格式为 `.xlsx`（不支持旧版 `.xls`）。

**Q：生成的文件名包含非法字符**
> Windows 文件名不允许 `\ / : * ? " < > |` 这些字符。如果变量值含有上述字符，程序会自动替换为 `_`。

---

## 许可证

MIT License © 2026 LZS

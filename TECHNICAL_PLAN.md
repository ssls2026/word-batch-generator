# Word 批量生成器 C# 重构 - 技术方案与任务清单

## 📋 项目目标

将 Python 版本（4089 行代码）完整迁移到 C# WPF，确保：
- ✅ **功能 100% 对标**（25 个核心功能 + 15 个增强功能）
- ✅ **现代化 UI**（Material Design + 流畅动画）
- ✅ **性能优势**（启动 < 1 秒，体积 60-80MB）
- ✅ **代码质量**（SOLID 原则、完整注释、零内存泄漏）

---

## 🎯 核心功能对标清单

### 一、创建新方案页面（NewSchemePanel）

#### 已实现 ✅
- [x] Word 模板导入（OpenFileDialog）
- [x] 方案命名输入
- [x] 变量列表显示
- [x] 方案保存到 SchemeManager

#### 待实现 🔨
- [ ] **高保真 Word 预览**（HTML + WebView2）
  - [ ] 段落格式还原（对齐、缩进、行距）
  - [ ] Run 格式还原（字体、颜色、粗体、斜体）
  - [ ] 表格渲染（合并单元格）
  - [ ] 35px 页面边距模拟
  
- [ ] **鼠标选区变量标注**（核心难点）
  - [ ] JavaScript 选区检测（WebView2.ExecuteScriptAsync）
  - [ ] 单段落替换
  - [ ] 跨段落合并替换
  - [ ] Run 级精确格式保留
  - [ ] 模糊匹配（空白不敏感）
  
- [ ] **变量管理交互**
  - [ ] 变量点击定位（高亮+滚动）
  - [ ] 循环查找（再次点击跳到下一个）
  - [ ] 下拉框历史记录
  - [ ] 变量复用（关联到已有变量）
  
- [ ] **预览优化**
  - [ ] 滚动位置保持
  - [ ] 变量占位符高亮（紫色背景）
  - [ ] 实时状态反馈
  
- [ ] **方案编辑**
  - [ ] 加载已有方案（从 SchemeListPanel 触发）
  - [ ] 回显变量列表
  - [ ] 保留原配置

---

### 二、批量生成页面（GeneratePanel）

#### 已实现 ✅
- [x] 方案选择下拉框
- [x] Excel 文件导入
- [x] 输出目录选择
- [x] 文件名模板配置
- [x] 批量生成按钮
- [x] 进度条显示

#### 待实现 🔨
- [ ] **方案加载与恢复**
  - [ ] 加载 config.json 和 template_marked.docx
  - [ ] 显示方案信息（变量数量、模板名）
  - [ ] 恢复上次配置（输出目录、命名规则）
  - [ ] 跨电脑路径自适应
  
- [ ] **Excel 数据处理**
  - [ ] 导出 Excel 模板（美化表头、自动列宽）
  - [ ] 一键打开 Excel 文件
  - [ ] 表头缺失警告
  - [ ] 日期格式转换
  - [ ] 数据预览表格
  
- [ ] **手动粘贴输入方式** 【重要】
  - [ ] 动态生成变量输入框（Grid 布局）
  - [ ] PlainTextEdit（强制纯文本）
  - [ ] 多行数据解析
  - [ ] 固定值模式（CheckBox + 单行限制）
  - [ ] 固定值视觉样式（浅蓝背景+虚线边框）
  - [ ] 配置自动保存
  
- [ ] **双页签数据隔离**
  - [ ] TabControl（Excel / 手动粘贴）
  - [ ] 数据源切换逻辑
  - [ ] 记忆上次页签
  
- [ ] **高级配置**
  - [ ] 文件命名规则下拉框（历史记录）
  - [ ] 子目录分类功能（CheckBox + ComboBox）
  - [ ] 配置实时持久化
  
- [ ] **批量生成引擎**
  - [ ] DocumentFormat.OpenXml 变量替换
  - [ ] 文件名清洗（非法字符、长度限制）
  - [ ] 子目录按变量值创建
  - [ ] 容错处理（单行失败继续）
  - [ ] 成功统计
  
- [ ] **生成结果展示**
  - [ ] 文件卡片列表（ItemsControl）
  - [ ] "查看文档"按钮（Process.Start）
  - [ ] "定位"按钮（explorer /select）
  - [ ] "打开导出目录"按钮

---

### 三、方案管理页面（SchemeListPanel）

#### 已实现 ✅
- [x] 方案列表显示（ItemsControl + DataTemplate）
- [x] 使用方案按钮（跳转到 GeneratePanel）
- [x] 导出方案按钮（.wsp ZIP）
- [x] 删除方案按钮（确认弹窗）

#### 待实现 🔨
- [ ] **卡片优化**
  - [ ] 显示原始模板名
  - [ ] 前 5 个变量名（超过显示省略号）
  - [ ] 文件夹图标
  
- [ ] **导入方案包** 【重要】
  - [ ] 右上角导入按钮
  - [ ] .wsp 文件选择
  - [ ] ZIP 解压验证
  - [ ] 同名覆盖确认
  - [ ] 自动刷新列表
  
- [ ] **编辑方案**
  - [ ] 跳转到 NewSchemePanel
  - [ ] 传递方案名参数

---

### 四、主窗口与导航

#### 已实现 ✅
- [x] 左右分栏布局
- [x] 侧边栏导航（3 个 RadioButton）
- [x] Frame 页面切换
- [x] 配色方案（暖色调）

#### 待实现 🔨
- [ ] **导航优化**
  - [ ] 左侧指示条动画（Storyboard）
  - [ ] 页面切换动画（Fade + Slide）
  - [ ] Logo 矢量图标
  
- [ ] **启动恢复**
  - [ ] 读取 last_state.json
  - [ ] 自动加载上次方案
  - [ ] 直接进入批量生成页面

---

## 🎨 现代化 UI 设计规范

### 1. 配色方案（暖色调 Material Design）

```csharp
// 主题色
Primary: #C4612F      // 陶土橙（按钮、强调）
PrimaryLight: #D97943 // 浅橙（悬停）
PrimaryDark: #A94E22  // 深橙（按下）

// 背景色
Background: #F7F4EF   // 暖奶油色
Surface: #FBF9F5      // 浅白色
Card: #FFFFFF         // 纯白卡片

// 文字色
TextPrimary: #0F172A  // 深灰
TextSecondary: #475569 // 中灰
TextHint: #94A3B8     // 浅灰

// 边框色
Border: #E7E1D7       // 暖灰
Divider: #E2E8F0      // 分隔线

// 导航栏
NavBackground: #1F2421 // 深炭色
NavText: #CBD5E1       // 浅灰
NavActive: #F8FAFC     // 白色

// 功能色
Success: #10B981      // 绿色
Warning: #F59E0B      // 琥珀色
Error: #EF4444        // 红色
Info: #3B82F6         // 蓝色
```

### 2. 排版规范

```csharp
// 字体
FontFamily: "Microsoft YaHei UI", "微软雅黑", "Segoe UI"
FontWeightLight: 300
FontWeightRegular: 400
FontWeightSemiBold: 600
FontWeightBold: 700

// 字号
Display: 28px         // 页面标题
Headline: 20px        // 区块标题
Title: 16px           // 卡片标题
Body: 14px            // 正文
Caption: 12px         // 辅助说明

// 圆角
BorderRadiusSmall: 4px   // 输入框
BorderRadius: 8px        // 卡片、按钮
BorderRadiusLarge: 12px  // 大卡片
```

### 3. 间距系统

```csharp
Space4: 4px
Space8: 8px
Space12: 12px
Space16: 16px
Space20: 20px
Space24: 24px
Space32: 32px
```

### 4. 阴影层级

```xml
<!-- Elevation 1 (卡片) -->
<DropShadowEffect BlurRadius="8" 
                  ShadowDepth="2" 
                  Opacity="0.1" 
                  Color="#000000"/>

<!-- Elevation 2 (悬停) -->
<DropShadowEffect BlurRadius="12" 
                  ShadowDepth="4" 
                  Opacity="0.15" 
                  Color="#000000"/>

<!-- Elevation 3 (弹窗) -->
<DropShadowEffect BlurRadius="24" 
                  ShadowDepth="8" 
                  Opacity="0.2" 
                  Color="#000000"/>
```

---

## 🎬 动画设计规范

### 1. 时长与缓动

```csharp
// 时长
DurationFast: 150ms      // 按钮反馈
DurationNormal: 250ms    // 页面切换
DurationSlow: 400ms      // 复杂动画

// 缓动函数
EaseOut: CubicEase(EasingMode.EaseOut)
EaseInOut: CubicEase(EasingMode.EaseInOut)
```

### 2. 按钮交互动画

```xml
<!-- 悬停：轻微上浮 + 阴影加深 -->
<Storyboard x:Key="ButtonHoverStoryboard">
    <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.Y)"
                     To="-2" Duration="0:0:0.15">
        <DoubleAnimation.EasingFunction>
            <CubicEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
    </DoubleAnimation>
</Storyboard>

<!-- 按下：缩放 0.95 -->
<Storyboard x:Key="ButtonPressStoryboard">
    <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(ScaleTransform.ScaleX)"
                     To="0.95" Duration="0:0:0.1"/>
    <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(ScaleTransform.ScaleY)"
                     To="0.95" Duration="0:0:0.1"/>
</Storyboard>
```

### 3. 页面切换动画

```csharp
// Fade + Slide
var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(150));
var slideOut = new DoubleAnimation(0, -50, TimeSpan.FromMilliseconds(150));
fadeOut.Completed += (s, e) => {
    // 切换内容
    var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(250));
    var slideIn = new DoubleAnimation(50, 0, TimeSpan.FromMilliseconds(250));
};
```

### 4. 进度条动画

```xml
<!-- 不确定进度：波浪动画 -->
<Storyboard RepeatBehavior="Forever">
    <DoubleAnimation Storyboard.TargetProperty="(Canvas.Left)"
                     From="-100" To="500" Duration="0:0:1.5"/>
</Storyboard>
```

### 5. 卡片入场动画

```csharp
// 错开入场（Stagger）
for (int i = 0; i < items.Count; i++) {
    var card = items[i];
    card.Opacity = 0;
    card.RenderTransform = new TranslateTransform(0, 20);
    
    var delay = i * 50; // 每个卡片延迟 50ms
    var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(250)) {
        BeginTime = TimeSpan.FromMilliseconds(delay)
    };
    var slideIn = new DoubleAnimation(20, 0, TimeSpan.FromMilliseconds(250)) {
        BeginTime = TimeSpan.FromMilliseconds(delay)
    };
}
```

---

## 💻 代码规范

### 1. 命名约定

```csharp
// PascalCase
public class WordParser { }
public void GenerateBatch() { }
public string TemplatePath { get; set; }

// camelCase
private string _templatePath;
private readonly ILogger _logger;

// 常量全大写
private const int MAX_FILE_NAME_LENGTH = 200;

// 事件
public event EventHandler<SchemeEventArgs> SchemeSaved;

// 异步方法
public async Task<List<string>> GenerateBatchAsync() { }
```

### 2. 注释规范

```csharp
/// <summary>
/// 批量生成 Word 文档
/// </summary>
/// <param name="templatePath">模板文件路径</param>
/// <param name="dataList">数据列表，每行为一个字典</param>
/// <param name="outputDir">输出目录</param>
/// <param name="fileNameTemplate">文件名模板（例如：授权书-{{公司名称}}.docx）</param>
/// <param name="progress">进度回调，参数为 (当前进度, 总数)</param>
/// <returns>成功生成的文件数量</returns>
/// <exception cref="FileNotFoundException">模板文件不存在</exception>
/// <exception cref="ArgumentException">数据列表为空</exception>
public static async Task<int> GenerateBatchAsync(
    string templatePath,
    List<Dictionary<string, string>> dataList,
    string outputDir,
    string fileNameTemplate,
    IProgress<(int current, int total)>? progress = null)
{
    // 实现...
}
```

### 3. 资源释放

```csharp
// IDisposable 模式
public class WordParser : IDisposable
{
    private WordprocessingDocument? _document;
    private bool _disposed = false;

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                _document?.Dispose();
            }
            _disposed = true;
        }
    }
}

// using 语句
using var doc = WordprocessingDocument.Open(filePath, false);
```

### 4. 异步编程

```csharp
// 避免阻塞 UI 线程
private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
{
    BtnGenerate.IsEnabled = false;
    try
    {
        var progress = new Progress<(int, int)>((p) =>
        {
            ProgressBar.Value = (double)p.Item1 / p.Item2 * 100;
            TxtProgress.Text = $"正在生成: {p.Item1}/{p.Item2}";
        });

        await Task.Run(() => Generator.GenerateBatchAsync(
            templatePath, dataList, outputDir, fileNameTemplate, progress));
        
        ShowSuccess("生成完成！");
    }
    catch (Exception ex)
    {
        ShowError($"生成失败: {ex.Message}");
    }
    finally
    {
        BtnGenerate.IsEnabled = true;
    }
}
```

### 5. MVVM 模式（可选）

```csharp
// Model
public class Scheme
{
    public string Name { get; set; }
    public List<string> Variables { get; set; }
}

// ViewModel
public class SchemeListViewModel : INotifyPropertyChanged
{
    private ObservableCollection<Scheme> _schemes;
    public ObservableCollection<Scheme> Schemes
    {
        get => _schemes;
        set
        {
            _schemes = value;
            OnPropertyChanged();
        }
    }

    public ICommand LoadSchemeCommand { get; }
    public ICommand DeleteSchemeCommand { get; }
}

// View (XAML)
<ItemsControl ItemsSource="{Binding Schemes}">
    <ItemsControl.ItemTemplate>
        <DataTemplate>
            <Button Content="{Binding Name}" 
                    Command="{Binding DataContext.LoadSchemeCommand, 
                             RelativeSource={RelativeSource AncestorType=ItemsControl}}"
                    CommandParameter="{Binding}"/>
        </DataTemplate>
    </ItemsControl.ItemTemplate>
</ItemsControl>
```

---

## 🚧 待解决的技术难点

### 1. Word 预览中的选区检测 【最高优先级】

**Python 实现**：
- 使用 QTextEdit 的 textCursor() 获取选区
- QTextCursor 提供 selectionStart/End
- 通过 blockNumber 映射到 docx 段落

**C# 方案**：
- **方案 A**：WebView2 + JavaScript
  ```csharp
  // 注入 JS 监听选区
  await webView.CoreWebView2.ExecuteScriptAsync(@"
      document.addEventListener('mouseup', () => {
          const selection = window.getSelection();
          const text = selection.toString();
          const range = selection.getRangeAt(0);
          // 发送到 C#
          window.chrome.webview.postMessage({
              text: text,
              startOffset: range.startOffset,
              endOffset: range.endOffset
          });
      });
  ");
  
  // C# 接收
  webView.CoreWebView2.WebMessageReceived += (s, e) => {
      var json = e.WebMessageAsJson;
      // 解析并标注变量
  };
  ```
  - ✅ 优点：选区检测容易
  - ❌ 缺点：HTML 节点到 docx Run 的映射复杂

- **方案 B**：RichTextBox（弃用）
  - ❌ 不支持 HTML 渲染，格式还原困难

- **方案 C**：简化交互（推荐）
  - 用户手动输入起止文本
  - 自动模糊匹配定位
  - 降低实现复杂度

**推荐方案**：先实现 C（简化），后期优化为 A

### 2. 跨段落格式保留

**技术方案**：
```csharp
// 使用 Open XML SDK
var paragraphs = body.Elements<Paragraph>().ToList();
var startPara = paragraphs[startIndex];
var endPara = paragraphs[endIndex];

// 提取起点段落选区前的 Runs
var beforeRuns = startPara.Elements<Run>()
    .TakeWhile(r => GetRunText(r) != selectedText);

// 创建新段落合并
var newPara = new Paragraph();
newPara.Append(beforeRuns.Select(r => r.CloneNode(true)));
newPara.Append(new Run(new Text($"{{{{{variableName}}}}}")));

// 删除旧段落
for (int i = startIndex; i <= endIndex; i++)
{
    paragraphs[i].Remove();
}
```

### 3. 固定值模式的输入框样式

**技术方案**：
```xml
<Style x:Key="FixedValueTextBoxStyle" TargetType="TextBox">
    <Setter Property="Background" Value="#EFF6FF"/>
    <Setter Property="BorderBrush" Value="#93C5FD"/>
    <Setter Property="BorderThickness" Value="2"/>
    <Style.Triggers>
        <DataTrigger Binding="{Binding IsFixed}" Value="True">
            <Setter Property="BorderDashArray" Value="4 2"/>
        </DataTrigger>
    </Style.Triggers>
</Style>
```

---

## 📅 开发计划

### Phase 1: 核心功能（1 周）
- [ ] 完善 WordParser（HTML 渲染、变量替换）
- [ ] 完善 Generator（批量生成引擎）
- [ ] NewSchemePanel 基础功能（简化版变量标注）
- [ ] GeneratePanel Excel 导入导出
- [ ] 端到端流程打通

### Phase 2: 增强功能（1 周）
- [ ] 手动粘贴输入方式
- [ ] 固定值模式
- [ ] 双页签数据隔离
- [ ] 子目录分类
- [ ] 方案导入导出
- [ ] 配置自动保存

### Phase 3: UI 美化与动画（3 天）
- [ ] 实现所有动画效果
- [ ] 优化配色和排版
- [ ] 响应式布局
- [ ] Loading 状态
- [ ] Toast 通知

### Phase 4: 测试与优化（2 天）
- [ ] 功能测试
- [ ] 性能测试
- [ ] 内存泄漏检测
- [ ] 错误处理完善
- [ ] 用户文档

---

## 📝 立即行动

接下来我将：
1. 升级 WordParser 实现高保真 HTML 渲染
2. 优化 UI 主题和动画系统
3. 实现 NewSchemePanel 的完整交互
4. 实现 GeneratePanel 的双输入源
5. 完善所有细节功能

预计 **2 周**完成 100% 功能对标 + 现代化 UI！

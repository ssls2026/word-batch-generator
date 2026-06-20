using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using WordBatchGenerator.Core;
using System.Text.Json;

namespace WordBatchGenerator.Gui.Panels;

/// <summary>
/// 创建/编辑方案页面：导入 Word 模板、预览文档、标注变量、保存方案
/// </summary>
public partial class NewSchemePanel : Page
{
    private string _currentTemplatePath = string.Empty;
    private List<string> _variables = new();
    private WebViewSelectionMessage? _lastSelection;
    private bool _isEditing = false;
    private string _originalSchemeName = string.Empty;

    private class WebViewSelectionMessage
    {
        public string text { get; set; } = string.Empty;
        public int startPIndex { get; set; }
        public int endPIndex { get; set; }
        public int startOffset { get; set; }
        public int endOffset { get; set; }
    }

    public NewSchemePanel()
    {
        InitializeComponent();
        InitializeWebView();
    }

    /// <summary>
    /// 支持重新编辑方案的构造函数
    /// </summary>
    public NewSchemePanel(string schemeName) : this()
    {
        LoadSchemeForEdit(schemeName);
    }

    private System.Threading.Tasks.Task? _webViewInitTask;

    /// <summary>
    /// 初始化 WebView2 环境（异步完成）
    /// </summary>
    private void InitializeWebView()
    {
        _webViewInitTask = EnsureWebViewInitializedAsync();
    }

    private async System.Threading.Tasks.Task EnsureWebViewInitializedAsync()
    {
        try
        {
            await WordPreview.EnsureCoreWebView2Async(null);
            WordPreview.CoreWebView2.Settings.IsScriptEnabled = true;
            WordPreview.NavigationCompleted += WordPreview_NavigationCompleted;
            WordPreview.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

            if (string.IsNullOrEmpty(_currentTemplatePath))
            {
                Dispatcher.Invoke(() =>
                {
                    WordPreview.NavigateToString(GetWelcomeHtml());
                });
            }
        }
        catch (Exception)
        {
            // WebView2 runtime 未安装时降级到纯文本显示
        }
    }

    private string GetWelcomeHtml()
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html><head><meta charset='utf-8'>");
        sb.AppendLine("<style>");
        sb.AppendLine("body { font-family: 'Microsoft YaHei', sans-serif; background-color: #F7F4EF; color: #1F2421; margin: 0; padding: 20px; display: flex; justify-content: center; align-items: center; min-height: 100vh; box-sizing: border-box; }");
        sb.AppendLine("#paper-container { background-color: #ffffff; width: 100%; max-width: 680px; min-height: 480px; padding: 30px; box-sizing: border-box; border-radius: 12px; border: 1px dashed #D4CEC2; box-shadow: 0 4px 20px rgba(0,0,0,0.03); text-align: center; display: flex; flex-direction: column; justify-content: center; align-items: center; }");
        sb.AppendLine(".icon { font-size: 48px; margin-bottom: 16px; }");
        sb.AppendLine("h2 { margin: 0 0 8px 0; font-size: 20px; font-weight: bold; color: #C4612F; }");
        sb.AppendLine("p.subtitle { margin: 0 0 24px 0; font-size: 13px; color: #5C635D; max-width: 440px; line-height: 1.5; }");
        sb.AppendLine(".steps-container { display: flex; justify-content: space-between; align-items: stretch; width: 100%; max-width: 560px; margin-bottom: 24px; gap: 12px; }");
        sb.AppendLine(".step-card { flex: 1; background: #FBF9F5; border: 1px solid #EBE7DE; border-radius: 8px; padding: 16px 10px; display: flex; flex-direction: column; align-items: center; }");
        sb.AppendLine(".step-num { width: 20px; height: 20px; border-radius: 10px; background: #C4612F; color: #ffffff; display: flex; justify-content: center; align-items: center; font-size: 11px; font-weight: bold; margin-bottom: 8px; }");
        sb.AppendLine(".step-title { font-size: 12px; font-weight: bold; color: #1F2421; margin-bottom: 4px; }");
        sb.AppendLine(".step-desc { font-size: 11px; color: #5C635D; line-height: 1.4; }");
        sb.AppendLine(".preview-mock { border: 1px solid #EBE7DE; border-radius: 6px; padding: 10px 14px; width: 100%; max-width: 440px; background: #FAFAFA; text-align: left; font-size: 12px; line-height: 1.6; color: #5C635D; box-sizing: border-box; }");
        sb.AppendLine(".variable { font-weight: bold; color: #4E7A8A; background-color: #F0F5F7; border: 1px solid #C5D6DC; border-radius: 4px; padding: 1px 3px; }");
        sb.AppendLine("</style>");
        sb.AppendLine("</head><body>");
        sb.AppendLine("<div id='paper-container'>");
        sb.AppendLine("  <div class='icon'>📄</div>");
        sb.AppendLine("  <h2>Word 模板解析与实时标注</h2>");
        sb.AppendLine("  <p class='subtitle'>导入您的 Word 模板文档（.docx），在这里即可直接预览文档内容并划选文字定义变量，帮助您快速建立批量生成方案。</p>");
        sb.AppendLine("  <div class='steps-container'>");
        sb.AppendLine("    <div class='step-card'>");
        sb.AppendLine("      <div class='step-num'>1</div>");
        sb.AppendLine("      <div class='step-title'>导入 Word 模板</div>");
        sb.AppendLine("      <div class='step-desc'>点击上方“导入模板”按钮，载入您的文档。</div>");
        sb.AppendLine("    </div>");
        sb.AppendLine("    <div class='step-card'>");
        sb.AppendLine("      <div class='step-num'>2</div>");
        sb.AppendLine("      <div class='step-title'>划选文档内容</div>");
        sb.AppendLine("      <div class='step-desc'>用鼠标在预览区直接选中需要被替换的文字。</div>");
        sb.AppendLine("    </div>");
        sb.AppendLine("    <div class='step-card'>");
        sb.AppendLine("      <div class='step-num'>3</div>");
        sb.AppendLine("      <div class='step-title'>标注为变量</div>");
        sb.AppendLine("      <div class='step-desc'>在右侧输入变量名并点击“标注”，即可完成绑定。</div>");
        sb.AppendLine("    </div>");
        sb.AppendLine("  </div>");
        sb.AppendLine("  <div class='preview-mock'>");
        sb.AppendLine("    <strong>预览效果演示：</strong><br/>");
        sb.AppendLine("    尊敬的 <span class='variable'>{{客户姓名}}</span> 客户，您好！您的订单号为 <span class='variable'>{{订单编号}}</span>。");
        sb.AppendLine("  </div>");
        sb.AppendLine("</div>");
        sb.AppendLine("</body></html>");
        return sb.ToString();
    }

    /// <summary>
    /// WebView 消息接收处理
    /// </summary>
    private void CoreWebView2_WebMessageReceived(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var json = e.WebMessageAsJson;
            _lastSelection = JsonSerializer.Deserialize<WebViewSelectionMessage>(json);
            if (_lastSelection != null)
            {
                Dispatcher.Invoke(() =>
                {
                    TxtSelectionText.Text = $"已选中: \"{_lastSelection.text}\"";
                    TxtSelectionText.Foreground = (System.Windows.Media.Brush)Application.Current.Resources["PrimaryBrush"];
                    TxtSelectionText.FontWeight = FontWeights.Bold;
                });
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"解析选区消息失败: {ex.Message}");
        }
    }

    /// <summary>
    /// WebView 页面加载完毕后注入选区监听 JS
    /// </summary>
    private async void WordPreview_NavigationCompleted(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
    {
        if (WordPreview.CoreWebView2 == null) return;

        string js = @"
            function getCharacterOffsetWithin(parent, node, offset) {
                let currentOffset = 0;
                const walker = document.createTreeWalker(parent, NodeFilter.SHOW_TEXT, null, false);
                while (walker.nextNode()) {
                    const textNode = walker.currentNode;
                    if (textNode === node) {
                        currentOffset += offset;
                        return currentOffset;
                    }
                    currentOffset += textNode.textContent.length;
                }
                return currentOffset;
            }

            document.addEventListener('mouseup', () => {
                const selection = window.getSelection();
                if (!selection.rangeCount) return;
                const range = selection.getRangeAt(0);
                const text = selection.toString();
                if (!text) return;

                let startP = range.startContainer;
                while (startP && startP.nodeName !== 'P') {
                    startP = startP.parentNode;
                }
                let endP = range.endContainer;
                while (endP && endP.nodeName !== 'P') {
                    endP = endP.parentNode;
                }

                if (!startP || !endP) return;

                const startPIndex = parseInt(startP.getAttribute('data-p-index'));
                const endPIndex = parseInt(endP.getAttribute('data-p-index'));

                if (isNaN(startPIndex) || isNaN(endPIndex)) return;

                const startOffset = getCharacterOffsetWithin(startP, range.startContainer, range.startOffset);
                const endOffset = getCharacterOffsetWithin(endP, range.endContainer, range.endOffset);

                window.chrome.webview.postMessage({
                    text: text,
                    startPIndex: startPIndex,
                    endPIndex: endPIndex,
                    startOffset: startOffset,
                    endOffset: endOffset
                });
            });
        ";
        await WordPreview.CoreWebView2.ExecuteScriptAsync(js);
    }

    /// <summary>
    /// 浏览模板文件
    /// </summary>
    private void BtnBrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Word 文档 (*.docx)|*.docx",
            Title = "选择 Word 模板文件"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                // 拷贝至临时路径，不修改用户原始文件
                var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                File.Copy(dialog.FileName, tempPath, true);

                _currentTemplatePath = tempPath;
                TxtTemplatePath.Text = dialog.FileName; // 显示原文件名
                LoadPreview();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"载入文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private async void LoadPreview()
    {
        if (string.IsNullOrEmpty(_currentTemplatePath) || !File.Exists(_currentTemplatePath))
            return;

        try
        {
            if (_webViewInitTask != null)
            {
                await _webViewInitTask;
            }

            if (WordPreview.CoreWebView2 == null) return;

            var html = WordParser.ConvertToHtml(_currentTemplatePath);
            WordPreview.NavigateToString(html);

            TxtStatus.Text = "✅ 模板加载成功，可以在预览区选择文本标注变量";
            TxtStatus.Foreground = System.Windows.Media.Brushes.Green;
            StatusBar.Visibility = Visibility.Visible;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载预览失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 标注选区为变量
    /// </summary>
    private void BtnAddVariable_Click(object sender, RoutedEventArgs e)
    {
        if (_lastSelection == null)
        {
            MessageBox.Show("请先在 Word 文档预览区中划选需要作为变量的文字！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var variableName = TxtVariableName.Text.Trim();

        if (string.IsNullOrEmpty(variableName))
        {
            MessageBox.Show("请输入变量名称！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            // 执行 OpenXML 文字替换并合并
            WordParser.ReplaceTextAcrossParagraphs(
                _currentTemplatePath,
                _lastSelection.startPIndex,
                _lastSelection.endPIndex,
                _lastSelection.startOffset,
                _lastSelection.endOffset - 1,
                variableName
            );

            // 更新状态 (仅在是新变量时加入集合与列表显示)
            if (!_variables.Contains(variableName))
            {
                _variables.Add(variableName);
                UpdateVariableNameComboBox();
                LstVariables.Items.Add(variableName);
            }
            
            TxtVariableName.Text = string.Empty;
            UpdateVariableCount();

            // 重置选区
            _lastSelection = null;
            TxtSelectionText.Text = "请在左侧预览区中划选变量文本";
            TxtSelectionText.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#8A8070"));
            TxtSelectionText.FontWeight = FontWeights.Medium;

            // 重新载入预览
            LoadPreview();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"标注变量失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnDeleteVariableItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is string varName)
        {
            RemoveVariable(varName);
        }
    }

    private void RemoveVariable(string varName)
    {
        if (_variables.Contains(varName))
        {
            _variables.Remove(varName);
            UpdateVariableNameComboBox();
            LstVariables.Items.Remove(varName);
            UpdateVariableCount();
        }
    }

    private void UpdateVariableCount()
    {
        TxtVariableCount.Text = $"已标注 {_variables.Count} 个变量";
    }

    /// <summary>
    /// 保存方案
    /// </summary>
    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        var schemeName = TxtSchemeName.Text.Trim();

        if (string.IsNullOrEmpty(schemeName))
        {
            MessageBox.Show("请输入方案名称！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrEmpty(_currentTemplatePath))
        {
            MessageBox.Show("请先导入 Word 模板文件！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (_variables.Count == 0)
        {
            MessageBox.Show("请至少标注一个变量！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 验证非法字符
        var invalidChars = new char[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };
        if (schemeName.IndexOfAny(invalidChars) >= 0)
        {
            MessageBox.Show("方案名称包含非法字符，请重新输入！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            // 覆盖检测
            var schemesDir = SchemeManager.GetSchemesDirectory();
            var schemePath = Path.Combine(schemesDir, schemeName);
            if (Directory.Exists(schemePath) && (!_isEditing || _originalSchemeName != schemeName))
            {
                var result = MessageBox.Show($"已存在名为 '{schemeName}' 的方案。确认覆盖它吗？", "确认覆盖", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;
            }

            var scheme = new Scheme
            {
                Name = schemeName,
                TemplatePath = _currentTemplatePath, // 传递当前被修改的临时文档
                Variables = _variables,
                CreatedAt = DateTime.Now,
                UpdatedAt = DateTime.Now
            };

            // 如果是在编辑状态下且修改了方案名，需要把旧目录清理掉
            if (_isEditing && _originalSchemeName != schemeName)
            {
                SchemeManager.DeleteScheme(_originalSchemeName);
            }

            // 保存方案（会自动拷贝临时模板到方案目录中）
            SchemeManager.SaveScheme(scheme);
            
            // 保持原样，方便用户修改。更新为编辑状态
            _isEditing = true;
            _originalSchemeName = schemeName;
            _lastSelection = null;
            TxtSelectionText.Text = "当前选中: 未选择文本";

            // 刷新 ComboBox 下拉列表
            UpdateVariableNameComboBox();

            // 状态栏反馈
            TxtStatus.Text = "✅ 方案保存成功！可继续进行修改编辑";
            TxtStatus.Foreground = System.Windows.Media.Brushes.Green;
            StatusBar.Visibility = Visibility.Visible;
            
            MessageBox.Show("方案保存成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 加载已有方案进行重新编辑
    /// </summary>
    private void LoadSchemeForEdit(string schemeName)
    {
        try
        {
            var scheme = SchemeManager.LoadScheme(schemeName);
            if (scheme == null) return;

            _isEditing = true;
            _originalSchemeName = schemeName;

            TxtSchemeName.Text = scheme.Name;
            
            // 拷贝方案的 template.docx 至临时目录供编辑
            var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var schemeTemplate = Path.Combine(SchemeManager.GetSchemesDirectory(), schemeName, "template.docx");
            if (File.Exists(schemeTemplate))
            {
                File.Copy(schemeTemplate, tempPath, true);
                _currentTemplatePath = tempPath;
                TxtTemplatePath.Text = schemeTemplate;
            }

            _variables.Clear();
            LstVariables.Items.Clear();
            foreach (var varName in scheme.Variables)
            {
                _variables.Add(varName);
                LstVariables.Items.Add(varName);
            }

            UpdateVariableNameComboBox();
            UpdateVariableCount();
            LoadPreview();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载编辑方案失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private string _lastClickedVariable = "";
    private int _lastScrollIndex = -1;

    /// <summary>
    /// 点击或按键移动到列表项时，定位到 WebView 预览中的对应文本
    /// </summary>
    private void LstVariables_PreviewMouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var dependencyObject = (DependencyObject)e.OriginalSource;
        while (dependencyObject != null && dependencyObject != LstVariables)
        {
            if (dependencyObject is ListBoxItem item)
            {
                var varName = item.Content as string;
                if (varName != null)
                {
                    NavigateToVariable(varName);
                }
                break;
            }
            dependencyObject = System.Windows.Media.VisualTreeHelper.GetParent(dependencyObject);
        }
    }

    private void LstVariables_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Up || e.Key == System.Windows.Input.Key.Down)
        {
            if (LstVariables.SelectedItem is string varName)
            {
                if (varName != _lastClickedVariable)
                {
                    NavigateToVariable(varName);
                }
            }
        }
    }

    private void NavigateToVariable(string varName)
    {
        if (string.IsNullOrEmpty(varName)) return;

        if (_lastClickedVariable == varName)
        {
            _lastScrollIndex++;
        }
        else
        {
            _lastClickedVariable = varName;
            _lastScrollIndex = 0;
        }

        ScrollToVariableInWebView(varName, _lastScrollIndex);
    }

    private async void ScrollToVariableInWebView(string varName, int index)
    {
        if (WordPreview.CoreWebView2 == null) return;

        string varPlaceholder = "{{" + varName + "}}";
        string js = $@"
            (function() {{
                const spans = Array.from(document.querySelectorAll('span.variable'));
                const matches = spans.filter(span => span.textContent.trim() === '{varPlaceholder}' || span.textContent.trim() === '{varName}');
                if (matches.length > 0) {{
                    const target = matches[{index} % matches.length];
                    target.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                    
                    // Add the flash-active class instantly
                    target.classList.add('flash-active');
                    
                    // Force a reflow
                    target.offsetHeight;
                    
                    // Remove after 300ms to let CSS transitions smoothly fade out the highlight
                    setTimeout(() => {{
                        target.classList.remove('flash-active');
                    }}, 300);
                }}
            }})();
        ";
        await WordPreview.CoreWebView2.ExecuteScriptAsync(js);
    }

    private void UpdateVariableNameComboBox()
    {
        TxtVariableName.ItemsSource = null;
        TxtVariableName.ItemsSource = _variables;
    }
}

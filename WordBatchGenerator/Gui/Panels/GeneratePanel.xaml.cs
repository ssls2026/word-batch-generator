using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WordBatchGenerator.Core;

namespace WordBatchGenerator.Gui.Panels;

/// <summary>
/// 批量生成页面：加载方案、导入 Excel、手动粘贴文本、配置输出、批量生成 Word 文档
/// </summary>
public partial class GeneratePanel : Page
{
    private Scheme? _currentScheme;
    private List<Dictionary<string, string>> _excelData = new();
    private Dictionary<string, TextBox> _textBoxes = new();
    private Dictionary<string, CheckBox> _checkBoxes = new();
    private bool _isInitializing = false;

    public GeneratePanel()
    {
        InitializeComponent();
        LoadSchemes();
        LoadLastScheme();
    }

    /// <summary>
    /// 带参数构造函数（从方案列表页跳转过来）
    /// </summary>
    public GeneratePanel(string schemeName) : this()
    {
        LoadScheme(schemeName);
    }

    /// <summary>
    /// 加载方案名称到下拉框
    /// </summary>
    private void LoadSchemes()
    {
        var schemes = SchemeManager.GetAllSchemes();
        CmbSchemes.ItemsSource = schemes.Select(s => s.Name).ToList();
    }

    /// <summary>
    /// 自动恢复上次使用的方案
    /// </summary>
    private void LoadLastScheme()
    {
        var lastScheme = SchemeManager.GetLastScheme();
        if (!string.IsNullOrEmpty(lastScheme))
            CmbSchemes.SelectedItem = lastScheme;
    }

    /// <summary>
    /// 加载指定方案的配置
    /// </summary>
    private void LoadScheme(string schemeName)
    {
        _isInitializing = true;
        _currentScheme = SchemeManager.LoadScheme(schemeName);
        CmbSchemes.SelectedItem = schemeName;

        if (_currentScheme != null)
        {
            TxtOutputDir.Text = _currentScheme.DefaultOutputDir;
            TxtFileNameTemplate.Text = _currentScheme.FileNameTemplate;

            // 初始化子目录下拉变量
            CmbSubfolderVar.ItemsSource = _currentScheme.Variables;
            if (!string.IsNullOrEmpty(_currentScheme.DefaultSubfolderVar) && _currentScheme.Variables.Contains(_currentScheme.DefaultSubfolderVar))
            {
                ChkSubfolder.IsChecked = true;
                CmbSubfolderVar.SelectedItem = _currentScheme.DefaultSubfolderVar;
                CmbSubfolderVar.IsEnabled = true;
            }

            // 初始化插入变量下拉菜单
            CmbInsertVariable.SelectionChanged -= CmbInsertVariable_SelectionChanged;
            CmbInsertVariable.Items.Clear();
            CmbInsertVariable.Items.Add("[ 选择变量插入 ]");
            CmbInsertVariable.Items.Add("序号");
            foreach (var varName in _currentScheme.Variables)
            {
                CmbInsertVariable.Items.Add(varName);
            }
            CmbInsertVariable.SelectedIndex = 0;
            CmbInsertVariable.SelectionChanged += CmbInsertVariable_SelectionChanged;

            // 初始化手动粘贴动态控件
            SetupPasteTab();

            // 还原已保存的输入文本和复选框状态
            foreach (var kvp in _currentScheme.PastedTexts)
            {
                if (_textBoxes.TryGetValue(kvp.Key, out var tb))
                {
                    tb.Text = kvp.Value;
                }
            }

            foreach (var varName in _currentScheme.FixedVariables)
            {
                if (_checkBoxes.TryGetValue(varName, out var chk))
                {
                    chk.IsChecked = true;
                }
            }

            // 还原上次活动页签
            if (_currentScheme.LastTabIndex >= 0 && _currentScheme.LastTabIndex < DataInputTabs.Items.Count)
            {
                DataInputTabs.SelectedIndex = _currentScheme.LastTabIndex;
            }

            // 触发对应数据源的预览
            if (DataInputTabs.SelectedIndex == 1)
            {
                UpdatePasteDataAndPreview();
            }
            else
            {
                // Excel 模式重置
                TxtExcelPath.Clear();
                _excelData = new List<Dictionary<string, string>>();
                RefreshDataPreviewGrid(null);
                DataPreviewBorder.Visibility = Visibility.Collapsed;
                BtnOpenExcel.IsEnabled = false;
                BtnGenerate.IsEnabled = false;
            }

            PanelResults.Visibility = Visibility.Collapsed;
            ResultItemsControl.ItemsSource = null;
        }

        _isInitializing = false;
    }

    /// <summary>
    /// 方案下拉框选择变化
    /// </summary>
    private void CmbSchemes_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CmbSchemes.SelectedItem is string schemeName)
            LoadScheme(schemeName);
    }

    /// <summary>
    /// 双页签切换事件
    /// </summary>
    private void DataInputTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.Source is TabControl)
        {
            if (DataInputTabs.SelectedIndex == 0)
            {
                // 导入 Excel 页签
                if (_excelData != null && DataPreviewGrid.ItemsSource != _excelData && TxtExcelPath.Text.Length > 0)
                {
                    LoadExcelPreview(TxtExcelPath.Text);
                }
                else if (string.IsNullOrEmpty(TxtExcelPath.Text))
                {
                    _excelData = new List<Dictionary<string, string>>();
                    RefreshDataPreviewGrid(null);
                    DataPreviewBorder.Visibility = Visibility.Collapsed;
                    BtnGenerate.IsEnabled = false;
                }
            }
            else if (DataInputTabs.SelectedIndex == 1)
            {
                // 手动粘贴页签
                UpdatePasteDataAndPreview();
            }

            AutoSaveConfig();
        }
    }

    /// <summary>
    /// 浏览并导入 Excel 数据源
    /// </summary>
    private void BtnBrowseExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
            Title = "选择 Excel 数据源"
        };

        if (dialog.ShowDialog() == true)
        {
            TxtExcelPath.Text = dialog.FileName;
            LoadExcelPreview(dialog.FileName);
            BtnOpenExcel.IsEnabled = true;
        }
    }

    /// <summary>
    /// 一键打开 Excel 文件
    /// </summary>
    private void BtnOpenExcel_Click(object sender, RoutedEventArgs e)
    {
        var path = TxtExcelPath.Text;
        if (File.Exists(path))
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开 Excel 文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    /// <summary>
    /// 导出 Excel 模板
    /// </summary>
    private void BtnExportExcelTemplate_Click(object sender, RoutedEventArgs e)
    {
        if (_currentScheme == null)
        {
            MessageBox.Show("请先选择方案！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel 文件 (*.xlsx)|*.xlsx",
            FileName = $"{_currentScheme.Name}_数据录入模板.xlsx",
            Title = "保存 Excel 模板"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                ExcelHandler.GenerateExcelTemplate(_currentScheme.Variables, dialog.FileName);
                MessageBox.Show("模板已生成！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{dialog.FileName}\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成模板失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    /// <summary>
    /// 文件名模板使用说明
    /// </summary>
    private void BtnFileNameHelp_Click(object sender, RoutedEventArgs e)
    {
        TxtFileNameTemplate.Text = "{{序号}}.docx";
        AutoSaveConfig();
    }

    /// <summary>
    /// 加载 Excel 数据预览
    /// </summary>
    private void LoadExcelPreview(string filePath)
    {
        try
        {
            _excelData = ExcelHandler.ReadExcel(filePath);

            RefreshDataPreviewGrid(_excelData);
            DataPreviewBorder.Visibility = Visibility.Visible;
            BtnGenerate.IsEnabled = _excelData.Count > 0;

            MessageBox.Show($"成功加载数据，共 {_excelData.Count} 行", "提示",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"读取 Excel 失败: {ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 选择输出目录（WPF 原生文件夹选择器）
    /// </summary>
    private void BtnBrowseOutputDir_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog
        {
            Title = "选择输出目录",
            Multiselect = false
        };

        if (dialog.ShowDialog() == true)
        {
            TxtOutputDir.Text = dialog.FolderName;
            AutoSaveConfig();
        }
    }

    /// <summary>
    /// 开始批量生成 Word 文档
    /// </summary>
    private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
    {
        // 参数校验
        if (_currentScheme == null)
        {
            MessageBox.Show("请先选择方案！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (DataInputTabs.SelectedIndex == 1)
        {
            // 粘贴源需要实时解析
            UpdatePasteDataAndPreview();
        }

        if (_excelData.Count == 0)
        {
            MessageBox.Show("请确保有需要生成的数据行！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrEmpty(TxtOutputDir.Text))
        {
            MessageBox.Show("请选择输出目录！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 验证数据列完整性
        var headers = _excelData.First().Keys.ToList();
        var missingVariables = Generator.ValidateData(_currentScheme.Variables, headers);
        if (missingVariables.Count > 0)
        {
            var missing = string.Join(", ", missingVariables);
            MessageBox.Show($"数据源中缺少以下变量列: {missing}", "数据不完整",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 显示进度 UI
        PanelProgress.Visibility = Visibility.Visible;
        PanelResults.Visibility = Visibility.Collapsed;
        BtnGenerate.IsEnabled = false;

        try
        {
            var templatePath = Path.Combine(
                SchemeManager.GetSchemesDirectory(),
                _currentScheme.Name,
                "template.docx"
            );

            var outputDir = TxtOutputDir.Text;
            var fileNameTemplate = TxtFileNameTemplate.Text;
            var subfolderVar = ChkSubfolder.IsChecked == true ? (CmbSubfolderVar.SelectedItem as string ?? "") : "";

            await Task.Run(() =>
            {
                Generator.GenerateBatch(
                    templatePath,
                    _excelData,
                    outputDir,
                    fileNameTemplate,
                    subfolderVar,
                    (current, total) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBar.Value = (double)current / total * 100;
                            TxtProgress.Text = $"正在生成: {current}/{total}";
                        });
                    }
                );
            });

            MessageBox.Show($"批量生成完成！共生成 {_excelData.Count} 个文件", "成功",
                MessageBoxButton.OK, MessageBoxImage.Information);

            // 渲染生成结果
            var generatedFiles = new List<string>();
            foreach (var data in _excelData)
            {
                var rawFileName = Generator.ReplaceVariables(fileNameTemplate, data);
                var fileName = Generator.SanitizeFileName(rawFileName);

                if (!string.IsNullOrEmpty(subfolderVar) && data.TryGetValue(subfolderVar, out var subDirValue) && !string.IsNullOrEmpty(subDirValue))
                {
                    var subfolder = Generator.SanitizeFileName(subDirValue);
                    generatedFiles.Add(Path.Combine(subfolder, fileName));
                }
                else
                {
                    generatedFiles.Add(fileName);
                }
            }

            ResultItemsControl.ItemsSource = generatedFiles;
            PanelResults.Visibility = Visibility.Visible;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"生成失败: {ex.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            PanelProgress.Visibility = Visibility.Collapsed;
            BtnGenerate.IsEnabled = true;
            ProgressBar.Value = 0;
        }
    }

    // ================= 手动粘贴处理逻辑 =================

    /// <summary>
    /// 动态生成手动粘贴面板变量文本框
    /// </summary>
    private void SetupPasteTab()
    {
        if (_currentScheme == null) return;

        PasteInputsGrid.Children.Clear();
        _textBoxes.Clear();
        _checkBoxes.Clear();

        foreach (var varName in _currentScheme.Variables)
        {
            var stackPanel = new StackPanel { Margin = new Thickness(0, 0, 16, 12) };

            var headerDock = new DockPanel { LastChildFill = false, Margin = new Thickness(0, 0, 0, 4) };

            var lbl = new TextBlock
            {
                Text = $"变量: {varName}",
                FontWeight = FontWeights.Bold,
                Foreground = (Brush)Application.Current.Resources["TextPrimaryBrush"],
                VerticalAlignment = VerticalAlignment.Center
            };
            DockPanel.SetDock(lbl, Dock.Left);
            headerDock.Children.Add(lbl);

            var chk = new CheckBox
            {
                Content = "固定相同值",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0),
                Tag = varName
            };
            DockPanel.SetDock(chk, Dock.Right);
            headerDock.Children.Add(chk);

            var tb = new TextBox
            {
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                Height = 90,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Tag = varName,
                ToolTip = "每行对应一个 Word 文件"
            };

            chk.Checked += (s, e) => OnFixedToggled(chk, tb, varName);
            chk.Unchecked += (s, e) => OnFixedToggled(chk, tb, varName);
            tb.TextChanged += (s, e) => OnPasteTextChanged(tb, varName);

            stackPanel.Children.Add(headerDock);
            stackPanel.Children.Add(tb);

            PasteInputsGrid.Children.Add(stackPanel);

            _textBoxes[varName] = tb;
            _checkBoxes[varName] = chk;
        }
    }

    /// <summary>
    /// 处理“固定相同值”复选框的切换事件
    /// </summary>
    private void OnFixedToggled(CheckBox chk, TextBox tb, string varName)
    {
        if (chk.IsChecked == true)
        {
            tb.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EFF6FF"));
            tb.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#93C5FD"));
            tb.BorderThickness = new Thickness(2);
            tb.ToolTip = "此变量在所有生成的文档中将保持相同值";

            // 强制截取第一行
            var lines = tb.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length > 1)
            {
                tb.Text = lines[0];
            }
        }
        else
        {
            tb.ClearValue(Control.BackgroundProperty);
            tb.ClearValue(Control.BorderBrushProperty);
            tb.ClearValue(Control.BorderThicknessProperty);
            tb.ToolTip = "每行对应一个 Word 文件";
        }

        UpdatePasteDataAndPreview();
        AutoSaveConfig();
    }

    /// <summary>
    /// 手动粘贴文本框字符改变事件
    /// </summary>
    private void OnPasteTextChanged(TextBox tb, string varName)
    {
        if (_checkBoxes.TryGetValue(varName, out var chk) && chk.IsChecked == true)
        {
            // 固定相同值，剥除换行
            var cleanText = tb.Text.Replace("\n", "").Replace("\r", "");
            if (cleanText != tb.Text)
            {
                var caretIndex = tb.CaretIndex;
                tb.Text = cleanText;
                tb.CaretIndex = Math.Min(caretIndex, tb.Text.Length);
            }
        }

        UpdatePasteDataAndPreview();
        AutoSaveConfig();
    }

    /// <summary>
    /// 触发手动粘贴数据解析并渲染到表格
    /// </summary>
    private void UpdatePasteDataAndPreview()
    {
        if (DataInputTabs.SelectedIndex != 1) return;

        var parsed = ParsePastedData();
        _excelData = parsed;

        if (parsed.Count > 0)
        {
            RefreshDataPreviewGrid(parsed);
            DataPreviewBorder.Visibility = Visibility.Visible;
            BtnGenerate.IsEnabled = true;
        }
        else
        {
            RefreshDataPreviewGrid(null);
            DataPreviewBorder.Visibility = Visibility.Collapsed;
            BtnGenerate.IsEnabled = false;
        }
    }

    /// <summary>
    /// 手动粘贴的数据流解析成 List grid
    /// </summary>
    private List<Dictionary<string, string>> ParsePastedData()
    {
        if (_currentScheme == null) return new List<Dictionary<string, string>>();

        var variables = _currentScheme.Variables;
        var varLines = new Dictionary<string, string[]>();
        int maxLines = 0;

        foreach (var varName in variables)
        {
            if (_textBoxes.TryGetValue(varName, out var tb))
            {
                var chk = _checkBoxes[varName];
                var isFixed = chk.IsChecked == true;

                var lines = tb.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None)
                                   .Select(l => l.Trim())
                                   .ToArray();

                // 除去最后的空行
                int lastNonEmpty = lines.Length - 1;
                while (lastNonEmpty >= 0 && string.IsNullOrEmpty(lines[lastNonEmpty]))
                {
                    lastNonEmpty--;
                }
                var cleanedLines = lines.Take(lastNonEmpty + 1).ToArray();
                varLines[varName] = cleanedLines;

                if (!isFixed)
                {
                    if (cleanedLines.Length > maxLines)
                    {
                        maxLines = cleanedLines.Length;
                    }
                }
            }
        }

        var parsedData = new List<Dictionary<string, string>>();

        if (maxLines == 0)
        {
            var rowDict = new Dictionary<string, string>();
            bool hasValue = false;
            foreach (var varName in variables)
            {
                var chk = _checkBoxes[varName];
                if (chk.IsChecked == true && varLines.TryGetValue(varName, out var lines) && lines.Length > 0)
                {
                    var val = lines[0];
                    rowDict[varName] = val;
                    if (!string.IsNullOrEmpty(val))
                        hasValue = true;
                }
                else
                {
                    rowDict[varName] = string.Empty;
                }
            }
            if (hasValue)
            {
                parsedData.Add(rowDict);
            }
            return parsedData;
        }

        for (int i = 0; i < maxLines; i++)
        {
            var rowDict = new Dictionary<string, string>();
            bool hasDynamicValue = false;

            foreach (var varName in variables)
            {
                var chk = _checkBoxes[varName];
                var isFixed = chk.IsChecked == true;

                if (varLines.TryGetValue(varName, out var lines))
                {
                    if (isFixed)
                    {
                        var val = lines.Length > 0 ? lines[0] : string.Empty;
                        rowDict[varName] = val;
                    }
                    else
                    {
                        var val = i < lines.Length ? lines[i] : string.Empty;
                        rowDict[varName] = val;
                        if (!string.IsNullOrEmpty(val))
                            hasDynamicValue = true;
                    }
                }
                else
                {
                    rowDict[varName] = string.Empty;
                }
            }

            if (hasDynamicValue)
            {
                parsedData.Add(rowDict);
            }
        }

        return parsedData;
    }

    // ================= 方案配置自动保存与恢复 =================

    /// <summary>
    /// 自动将输入配置存入 config.json
    /// </summary>
    private void AutoSaveConfig()
    {
        if (_isInitializing || _currentScheme == null) return;

        _currentScheme.DefaultOutputDir = TxtOutputDir.Text;
        _currentScheme.FileNameTemplate = TxtFileNameTemplate.Text;
        _currentScheme.DefaultSubfolderVar = ChkSubfolder.IsChecked == true ? (CmbSubfolderVar.SelectedItem as string ?? string.Empty) : string.Empty;
        _currentScheme.LastTabIndex = DataInputTabs.SelectedIndex;

        _currentScheme.PastedTexts.Clear();
        foreach (var kvp in _textBoxes)
        {
            _currentScheme.PastedTexts[kvp.Key] = kvp.Value.Text;
        }

        _currentScheme.FixedVariables.Clear();
        foreach (var kvp in _checkBoxes)
        {
            if (kvp.Value.IsChecked == true)
            {
                _currentScheme.FixedVariables.Add(kvp.Key);
            }
        }

        try
        {
            SchemeManager.SaveScheme(_currentScheme);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"自动保存配置失败: {ex.Message}");
        }
    }

    private void ChkSubfolder_Changed(object sender, RoutedEventArgs e)
    {
        if (CmbSubfolderVar != null)
        {
            CmbSubfolderVar.IsEnabled = ChkSubfolder.IsChecked == true;
        }
        AutoSaveConfig();
    }

    private void CmbSubfolderVar_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        AutoSaveConfig();
    }

    private void BtnFileNameHelp_TextChanged(object sender, TextChangedEventArgs e)
    {
        AutoSaveConfig();
    }

    // ================= 生成结果拉起与定位 =================

    private void BtnOpenResult_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string relativePath)
        {
            var fullPath = Path.Combine(TxtOutputDir.Text, relativePath);
            if (File.Exists(fullPath))
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = fullPath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"打开文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private void CmbInsertVariable_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CmbInsertVariable.SelectedIndex > 0)
        {
            var selectedVar = CmbInsertVariable.SelectedItem as string;
            if (!string.IsNullOrEmpty(selectedVar))
            {
                string insertText = "{{" + selectedVar + "}}";
                string currentText = TxtFileNameTemplate.Text;

                if (string.IsNullOrEmpty(currentText) || currentText == "文档-{{序号}}" || currentText == "文档-{{序号}}.docx")
                {
                    TxtFileNameTemplate.Text = insertText + ".docx";
                }
                else
                {
                    string baseText = currentText;
                    if (baseText.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        baseText = baseText.Substring(0, baseText.Length - 5);
                    }

                    if (string.IsNullOrEmpty(baseText))
                    {
                        TxtFileNameTemplate.Text = insertText + ".docx";
                    }
                    else
                    {
                        TxtFileNameTemplate.Text = baseText + "-" + insertText + ".docx";
                    }
                }
                AutoSaveConfig();
            }

            // 重置回默认占位选项
            CmbInsertVariable.SelectedIndex = 0;
        }
    }

    private void BtnLocateResult_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string relativePath)
        {
            var fullPath = Path.Combine(TxtOutputDir.Text, relativePath);
            if (File.Exists(fullPath))
            {
                try
                {
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{fullPath}\"");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"定位文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    /// <summary>
    /// 动态刷新数据预览表格的列和数据源，解决 Dictionary 绑定只读 Count/Keys 等 metadata 列的 WPF Bug
    /// </summary>
    private void RefreshDataPreviewGrid(System.Collections.Generic.List<System.Collections.Generic.Dictionary<string, string>>? data)
    {
        DataPreviewGrid.ItemsSource = null;
        DataPreviewGrid.Columns.Clear();

        if (data == null || data.Count == 0)
        {
            if (TxtDataPreviewSubtitle != null)
            {
                TxtDataPreviewSubtitle.Text = "(共 0 行)";
            }
            return;
        }

        if (TxtDataPreviewSubtitle != null)
        {
            TxtDataPreviewSubtitle.Text = $"(共 {data.Count} 行)";
        }

        // Clone list and inject "序号"
        var previewList = new List<Dictionary<string, string>>();
        for (int i = 0; i < data.Count; i++)
        {
            var newDict = new Dictionary<string, string>
            {
                ["序号"] = (i + 1).ToString()
            };
            foreach (var kvp in data[i])
            {
                newDict[kvp.Key] = kvp.Value;
            }
            previewList.Add(newDict);
        }

        // 1. Add "序号" column (Fixed width 60, centered content & header)
        var indexCol = new DataGridTextColumn
        {
            Header = "序号",
            Binding = new System.Windows.Data.Binding("[序号]"),
            Width = new DataGridLength(60)
        };

        var centerTextBlockStyle = new Style(typeof(TextBlock));
        centerTextBlockStyle.Setters.Add(new Setter(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center));
        centerTextBlockStyle.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
        indexCol.ElementStyle = centerTextBlockStyle;

        var centerHeaderStyle = new Style(typeof(System.Windows.Controls.Primitives.DataGridColumnHeader));
        centerHeaderStyle.Setters.Add(new Setter(System.Windows.Controls.Primitives.DataGridColumnHeader.HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
        if (DataPreviewGrid.TryFindResource(typeof(System.Windows.Controls.Primitives.DataGridColumnHeader)) is Style baseHeaderStyle)
        {
            centerHeaderStyle.BasedOn = baseHeaderStyle;
        }
        indexCol.HeaderStyle = centerHeaderStyle;

        DataPreviewGrid.Columns.Add(indexCol);

        // 2. Add other variable columns
        var keys = data.SelectMany(d => d.Keys).Distinct().Where(k => k != "序号").ToList();

        foreach (var key in keys)
        {
            var column = new DataGridTextColumn
            {
                Header = key,
                Binding = new System.Windows.Data.Binding($"[{key}]"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            };
            DataPreviewGrid.Columns.Add(column);
        }

        DataPreviewGrid.ItemsSource = previewList;
    }
}

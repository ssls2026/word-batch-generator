using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using WordBatchGenerator.Core;

namespace WordBatchGenerator.Gui.Panels;

/// <summary>
/// 方案管理页面：展示已有方案、导入导出、编辑、删除
/// </summary>
public partial class SchemeListPanel : Page
{
    public SchemeListPanel()
    {
        InitializeComponent();
        LoadSchemes();
        this.Loaded += (s, e) => LoadSchemes();
    }

    /// <summary>
    /// 加载方案列表并刷新 UI
    /// </summary>
    public void LoadSchemes()
    {
        var schemes = SchemeManager.GetAllSchemes();

        if (schemes.Count == 0)
        {
            TxtEmpty.Visibility = Visibility.Visible;
        }
        else
        {
            TxtEmpty.Visibility = Visibility.Collapsed;
        }

        SchemeItemsControl.ItemsSource = schemes;
    }

    /// <summary>
    /// 使用方案 — 跳转到批量生成页面
    /// </summary>
    private void BtnUseScheme_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string schemeName)
        {
            SchemeManager.SaveLastScheme(schemeName);
            NavigationService?.Navigate(new GeneratePanel(schemeName));
        }
    }

    /// <summary>
    /// 编辑方案 — 跳转到创建方案页面进行重新编辑
    /// </summary>
    private void BtnEditScheme_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string schemeName)
        {
            NavigationService?.Navigate(new NewSchemePanel(schemeName));
        }
    }

    /// <summary>
    /// 导出方案为 .wsp 压缩包
    /// </summary>
    private void BtnExportScheme_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string schemeName)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "方案包文件 (*.wsp)|*.wsp",
                FileName = $"{schemeName}.wsp",
                Title = "导出方案"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    SchemeManager.ExportScheme(schemeName, dialog.FileName);
                    MessageBox.Show("导出成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    /// <summary>
    /// 删除方案
    /// </summary>
    private void BtnDeleteScheme_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string schemeName)
        {
            var result = MessageBox.Show(
                $"确定要删除方案 '{schemeName}' 吗？此操作不可恢复。",
                "确认删除",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning
            );

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    SchemeManager.DeleteScheme(schemeName);
                    MessageBox.Show("删除成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    LoadSchemes();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"删除失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    /// <summary>
    /// 导入 .wsp 方案包
    /// </summary>
    private void BtnImportScheme_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "方案包文件 (*.wsp)|*.wsp",
            Title = "导入方案"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                var fileName = System.IO.Path.GetFileNameWithoutExtension(dialog.FileName);
                SchemeManager.ImportScheme(dialog.FileName, fileName);
                MessageBox.Show("导入成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                LoadSchemes();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

using System.Windows;
using System.Windows.Controls;
using WordBatchGenerator.Core;
using WordBatchGenerator.Gui.Panels;

namespace WordBatchGenerator;

/// <summary>
/// 主窗口：自定义标题栏 + 侧边导航 + 帧切换
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        // 标记起始页
        BtnNewScheme.Tag = new NewSchemePanel();
        BtnSchemeList.Tag = new SchemeListPanel();

        var lastScheme = SchemeManager.GetLastScheme();
        if (!string.IsNullOrEmpty(lastScheme))
        {
            BtnGenerate.Tag = new GeneratePanel(lastScheme);
            BtnGenerate.IsChecked = true;
            ContentFrame.Navigate((Page)BtnGenerate.Tag);
        }
        else
        {
            BtnGenerate.Tag = new GeneratePanel();
            BtnNewScheme.IsChecked = true;
            ContentFrame.Navigate((Page)BtnNewScheme.Tag);
        }
    }

    // ========== 窗口控制 ==========

    private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2) BtnMaximize_Click(sender, e);
        else if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed) DragMove();
    }

    private void BtnMinimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void BtnMaximize_Click(object sender, RoutedEventArgs e)
        => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;

    private void BtnClose_Click(object sender, RoutedEventArgs e) => Application.Current.Shutdown();

    // ========== 导航 ==========

    private void NavButton_Checked(object sender, RoutedEventArgs e)
    {
        if (sender is RadioButton btn && btn.Tag is Page page)
            ContentFrame.Navigate(page);
    }
}

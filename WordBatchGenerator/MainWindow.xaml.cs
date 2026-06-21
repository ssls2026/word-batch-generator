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

        // 加载历史应用状态与位置
        var state = SchemeManager.GetAppStateData();

        // 1. 恢复窗口位置和大小
        if (state.WindowWidth.HasValue && state.WindowHeight.HasValue)
        {
            // 限制不能小于设计的 MinWidth/MinHeight 且不能大于屏幕可用范围
            double screenWidth = SystemParameters.WorkArea.Width;
            double screenHeight = SystemParameters.WorkArea.Height;

            double width = Math.Max(MinWidth, Math.Min(state.WindowWidth.Value, screenWidth));
            double height = Math.Max(MinHeight, Math.Min(state.WindowHeight.Value, screenHeight));
            Width = width;
            Height = height;

            if (state.WindowLeft.HasValue && state.WindowTop.HasValue)
            {
                double left = Math.Max(0, Math.Min(state.WindowLeft.Value, screenWidth - width));
                double top = Math.Max(0, Math.Min(state.WindowTop.Value, screenHeight - height));
                Left = left;
                Top = top;
            }
        }

        if (state.IsMaximized)
        {
            WindowState = WindowState.Maximized;
        }

        // 2. 加载和导航到上次的页面
        var lastScheme = state.LastScheme;
        var lastPage = state.LastPage;

        BtnGenerate.Tag = !string.IsNullOrEmpty(lastScheme) ? new GeneratePanel(lastScheme) : new GeneratePanel();

        // 根据历史打开页面进行回显导航
        if (lastPage == "Generate")
        {
            BtnGenerate.IsChecked = true;
            ContentFrame.Navigate((Page)BtnGenerate.Tag);
        }
        else if (lastPage == "SchemeList")
        {
            BtnSchemeList.IsChecked = true;
            ContentFrame.Navigate((Page)BtnSchemeList.Tag);
        }
        else
        {
            BtnNewScheme.IsChecked = true;
            ContentFrame.Navigate((Page)BtnNewScheme.Tag);
        }

        // 监听关闭事件以保存位置与状态
        this.Closing += MainWindow_Closing;
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        try
        {
            var state = SchemeManager.GetAppStateData();

            // 记住当前活动的导航页面
            if (BtnGenerate.IsChecked == true)
                state.LastPage = "Generate";
            else if (BtnSchemeList.IsChecked == true)
                state.LastPage = "SchemeList";
            else
                state.LastPage = "NewScheme";

            // 记住窗口尺寸、位置与最大化状态
            if (WindowState == WindowState.Normal)
            {
                state.WindowLeft = Left;
                state.WindowTop = Top;
                state.WindowWidth = Width;
                state.WindowHeight = Height;
                state.IsMaximized = false;
            }
            else if (WindowState == WindowState.Maximized)
            {
                state.IsMaximized = true;
                // 最大化时，记录 RestoreBounds 的位置和大小，使下次还原为正常尺寸时符合预期
                state.WindowLeft = RestoreBounds.Left;
                state.WindowTop = RestoreBounds.Top;
                state.WindowWidth = RestoreBounds.Width;
                state.WindowHeight = RestoreBounds.Height;
            }

            SchemeManager.SaveAppStateData(state);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"关闭应用保存状态失败: {ex.Message}");
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

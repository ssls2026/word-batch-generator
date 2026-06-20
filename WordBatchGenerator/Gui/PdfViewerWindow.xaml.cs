using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using WordBatchGenerator.Core;

namespace WordBatchGenerator.Gui
{
    public partial class PdfViewerWindow : Window
    {
        private readonly string _pdfPath;

        public PdfViewerWindow(string pdfPath)
        {
            InitializeComponent();
            _pdfPath = pdfPath;
            TxtTitle.Text = $"📄 PDF 预览 - {Path.GetFileName(pdfPath)}";
            InitializeWebView();
        }

        private async void InitializeWebView()
        {
            try
            {
                // 创建专门的 WebView2 用户缓存目录
                var cacheDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordBatchGenerator",
                    "WebView2_Cache"
                );
                if (!Directory.Exists(cacheDir))
                {
                    Directory.CreateDirectory(cacheDir);
                }

                var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, cacheDir);
                await WebView.EnsureCoreWebView2Async(env);
                
                // 导航到本地 PDF 文件的 URL 路径
                WebView.Source = new Uri(_pdfPath);
                
                // 监听加载完毕显示页面
                WebView.CoreWebView2.DocumentTitleChanged += (s, e) => {
                    WebView.Visibility = Visibility.Visible;
                    PanelLoading.Visibility = Visibility.Collapsed;
                };
                
                // 双重保障：如果在 3 秒内未触发 Title 更改，也强制显示界面
                await System.Threading.Tasks.Task.Delay(3000);
                if (PanelLoading.Visibility == Visibility.Visible)
                {
                    WebView.Visibility = Visibility.Visible;
                    PanelLoading.Visibility = Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 PDF 预览引擎失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            var printWindow = new PrintWindow();
            printWindow.Owner = this;
            if (printWindow.ShowDialog() == true)
            {
                try
                {
                    this.Cursor = System.Windows.Input.Cursors.Wait;
                    Generator.PrintFileWithSettings(_pdfPath, printWindow.SelectedPrinter, printWindow.SelectedDuplex);
                    this.Cursor = System.Windows.Input.Cursors.Arrow;
                    MessageBox.Show("打印任务已成功发送！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    this.Cursor = System.Windows.Input.Cursors.Arrow;
                    MessageBox.Show($"打印文档失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }
    }
}

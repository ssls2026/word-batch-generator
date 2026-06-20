using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;

namespace WordBatchGenerator.Gui
{
    public partial class PrintWindow : Window
    {
        public string SelectedPrinter { get; private set; } = string.Empty;
        public string SelectedDuplex { get; private set; } = "onesided";
        public bool IsConfirmed { get; private set; } = false;

        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WordBatchGenerator",
            "print_settings.json"
        );

        public PrintWindow()
        {
            InitializeComponent();
            LoadPrinters();
            LoadDefaultSettings();
        }

        private void LoadPrinters()
        {
            try
            {
                var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters.Cast<string>().ToList();
                CmbPrinters.ItemsSource = printers;

                // 默认选择系统默认打印机
                var defaultPrinter = new System.Drawing.Printing.PrinterSettings().PrinterName;
                if (printers.Contains(defaultPrinter))
                {
                    CmbPrinters.SelectedItem = defaultPrinter;
                }
                else if (printers.Count > 0)
                {
                    CmbPrinters.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取打印机列表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDefaultSettings()
        {
            if (File.Exists(SettingsPath))
            {
                try
                {
                    var json = File.ReadAllText(SettingsPath);
                    var settings = JsonSerializer.Deserialize<PrintSettingsModel>(json);
                    if (settings != null)
                    {
                        if (CmbPrinters.ItemsSource is System.Collections.Generic.List<string> printers && printers.Contains(settings.DefaultPrinter))
                        {
                            CmbPrinters.SelectedItem = settings.DefaultPrinter;
                        }

                        switch (settings.DuplexMode.ToLower())
                        {
                            case "onesided":
                                RadOneSided.IsChecked = true;
                                break;
                            case "duplexlongedge":
                                RadDuplexLong.IsChecked = true;
                                break;
                            case "duplexshortedge":
                                RadDuplexShort.IsChecked = true;
                                break;
                        }

                        ChkSaveDefault.IsChecked = true;
                    }
                }
                catch { }
            }
        }

        private void SaveDefaultSettings()
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsPath);
                if (dir != null && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var settings = new PrintSettingsModel
                {
                    DefaultPrinter = SelectedPrinter,
                    DuplexMode = SelectedDuplex
                };

                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json);
            }
            catch { }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            if (CmbPrinters.SelectedItem == null)
            {
                MessageBox.Show("请选择一个打印机！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedPrinter = CmbPrinters.SelectedItem.ToString() ?? string.Empty;

            if (RadOneSided.IsChecked == true)
                SelectedDuplex = "onesided";
            else if (RadDuplexLong.IsChecked == true)
                SelectedDuplex = "duplexlongedge";
            else if (RadDuplexShort.IsChecked == true)
                SelectedDuplex = "duplexshortedge";

            if (ChkSaveDefault.IsChecked == true)
            {
                SaveDefaultSettings();
            }
            else
            {
                // 如果取消勾选，删除保存的默认设置文件
                if (File.Exists(SettingsPath))
                {
                    try { File.Delete(SettingsPath); } catch { }
                }
            }

            IsConfirmed = true;
            this.DialogResult = true;
            this.Close();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }
    }

    public class PrintSettingsModel
    {
        public string DefaultPrinter { get; set; } = string.Empty;
        public string DuplexMode { get; set; } = "onesided";
    }
}

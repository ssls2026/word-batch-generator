// 全局 using 别名 - 解决 WPF 和 WinForms 命名空间冲突
global using WpfMessageBox = System.Windows.MessageBox;
global using WpfButton = System.Windows.Controls.Button;
global using WpfRadioButton = System.Windows.Controls.RadioButton;
global using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
global using WpfSaveFileDialog = Microsoft.Win32.SaveFileDialog;
// Remove global using System.Windows.Forms to avoid type conflicts
// Use fully-qualified System.Windows.Forms.FolderBrowserDialog where needed
global using MessageBox = WordBatchGenerator.Gui.ModernMessageBox;

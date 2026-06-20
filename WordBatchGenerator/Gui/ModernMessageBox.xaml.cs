using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace WordBatchGenerator.Gui
{
    public partial class ModernMessageBox : Window
    {
        private MessageBoxResult _result = MessageBoxResult.None;
        private MessageBoxButton _buttonType;

        public ModernMessageBox()
        {
            InitializeComponent();
        }

        public static MessageBoxResult Show(string messageBoxText, string caption = "提示", MessageBoxButton button = MessageBoxButton.OK, MessageBoxImage icon = MessageBoxImage.None)
        {
            Window? activeWindow = Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive);
            if (activeWindow == null)
            {
                activeWindow = Application.Current.MainWindow;
            }
            return Show(activeWindow, messageBoxText, caption, button, icon);
        }

        public static MessageBoxResult Show(Window? owner, string messageBoxText, string caption = "提示", MessageBoxButton button = MessageBoxButton.OK, MessageBoxImage icon = MessageBoxImage.None)
        {
            var box = new ModernMessageBox();
            if (owner != null && owner.IsVisible)
            {
                box.Owner = owner;
            }
            
            box.TxtTitle.Text = caption;
            box.TxtMessage.Text = messageBoxText;
            box._buttonType = button;

            // Set Icon Emoji
            switch (icon)
            {
                case MessageBoxImage.Information:
                    box.TxtIcon.Text = "💡";
                    break;
                case MessageBoxImage.Warning:
                    box.TxtIcon.Text = "⚠️";
                    break;
                case MessageBoxImage.Error:
                    box.TxtIcon.Text = "❌";
                    break;
                case MessageBoxImage.Question:
                    box.TxtIcon.Text = "❓";
                    break;
                default:
                    box.TxtIcon.Text = "🔔";
                    break;
            }

            // Configure buttons based on MessageBoxButton
            switch (button)
            {
                case MessageBoxButton.OK:
                    box.BtnPrimary.Content = "确定";
                    box.BtnSecondary.Visibility = Visibility.Collapsed;
                    break;
                case MessageBoxButton.OKCancel:
                    box.BtnPrimary.Content = "确定";
                    box.BtnSecondary.Content = "取消";
                    box.BtnSecondary.Visibility = Visibility.Visible;
                    break;
                case MessageBoxButton.YesNo:
                    box.BtnPrimary.Content = "是";
                    box.BtnSecondary.Content = "否";
                    box.BtnSecondary.Visibility = Visibility.Visible;
                    break;
                case MessageBoxButton.YesNoCancel:
                    box.BtnPrimary.Content = "是";
                    box.BtnSecondary.Content = "否";
                    box.BtnSecondary.Visibility = Visibility.Visible;
                    break;
            }

            box.ShowDialog();
            return box._result;
        }

        private void BtnPrimary_Click(object sender, RoutedEventArgs e)
        {
            if (_buttonType == MessageBoxButton.YesNo || _buttonType == MessageBoxButton.YesNoCancel)
            {
                _result = MessageBoxResult.Yes;
            }
            else
            {
                _result = MessageBoxResult.OK;
            }
            this.DialogResult = true;
            this.Close();
        }

        private void BtnSecondary_Click(object sender, RoutedEventArgs e)
        {
            if (_buttonType == MessageBoxButton.YesNo || _buttonType == MessageBoxButton.YesNoCancel)
            {
                _result = MessageBoxResult.No;
            }
            else
            {
                _result = MessageBoxResult.Cancel;
            }
            this.DialogResult = false;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (_buttonType == MessageBoxButton.YesNo)
            {
                _result = MessageBoxResult.No;
            }
            else
            {
                _result = MessageBoxResult.Cancel;
            }
            this.DialogResult = false;
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
}

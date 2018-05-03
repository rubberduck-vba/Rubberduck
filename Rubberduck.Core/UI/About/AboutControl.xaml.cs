using System.Windows;
using System.Windows.Input;

namespace Rubberduck.UI.About
{
    /// <summary>
    /// Interaction logic for AboutControl.xaml
    /// </summary>
    public partial class AboutControl
    {
        public AboutControl()
        {
            InitializeComponent();
        }

        private void OnKeyDownHandler(object sender, KeyEventArgs e)
        {
            bool isControlCPressed = (Keyboard.IsKeyDown(Key.C) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)));
            if (isControlCPressed)
            {
                CopyVersionInfoToClipboard();
            }
        }

        private void CopyVersionInfo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            CopyVersionInfoToClipboard();
        }

        private void CopyVersionInfoToClipboard()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(Version.Text);
            sb.AppendLine(OperatingSystem.Text);
            sb.AppendLine(HostProduct.Text);
            sb.AppendLine(HostVersion.Text);
            sb.AppendLine(HostExecutable.Text);

            Clipboard.SetText(sb.ToString());
            System.Windows.MessageBox.Show(RubberduckUI.AboutWindow_CopyVersionMessage, RubberduckUI.AboutWindow_CopyVersionCaption);
        }
    }
}

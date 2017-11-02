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

        private void CopyVersionInfoToClipboard()
        {
            Clipboard.SetText(this.Version.Text);
            System.Windows.MessageBox.Show("Version information copied to clipboard.", "Copy successfull");
        }

        private void CopyVersionInfo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            CopyVersionInfoToClipboard();
        }
    }
}

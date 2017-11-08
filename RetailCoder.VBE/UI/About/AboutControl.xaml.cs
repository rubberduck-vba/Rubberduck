using System.Windows;
using System.Windows.Input;
using System;
using Application = System.Windows.Forms.Application;

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
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine($"Rubberduck version: {this.Version.Text}");
            sb.AppendLine($"Operating System: {Environment.OSVersion.VersionString}, {GetBitness(Environment.Is64BitOperatingSystem)}");
            sb.AppendLine($"Host Product: {Application.ProductName} {GetBitness(Environment.Is64BitProcess)}");
            sb.AppendLine($"Host Version: {Application.ProductVersion}");
            sb.AppendFormat($"Host Executable: {System.IO.Path.GetFileName(Application.ExecutablePath)}");

            Clipboard.SetText(sb.ToString());
            System.Windows.MessageBox.Show("Version information copied to clipboard.", "Copy successfull");

            string GetBitness(bool is64Bit)
            {
                if (is64Bit)
                {
                    return "x64";
                }
                else
                {
                    return "x86";
                }
            }
        }
    }
}

using System.Windows;
using System.Windows.Input;
using System;
using System.IO;
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
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Rubberduck version: {Version.Text}");
            sb.AppendLine($"Operating System: {Environment.OSVersion.VersionString}, {(Environment.Is64BitOperatingSystem ? "x64" : "x86")}");
            sb.AppendLine($"Host Product: {Application.ProductName} {(Environment.Is64BitProcess ? "x64" : "x86")}");
            sb.AppendLine($"Host Version: {Application.ProductVersion}");
            sb.AppendFormat($"Host Executable: {Path.GetFileName(Application.ExecutablePath).ToUpper()}"); // .ToUpper() used to convert ExceL.EXE -> EXCEL.EXE

            Clipboard.SetText(sb.ToString());
            System.Windows.MessageBox.Show(RubberduckUI.AboutWindow_CopyVersionMessage, RubberduckUI.AboutWindow_CopyVersionCaption);
        }
    }
}

using System;
using System.Diagnostics;
using NLog;
using NLog.Targets;
using Rubberduck.Resources.About;
using Rubberduck.UI.Command;
using Rubberduck.VersionCheck;
using Application = System.Windows.Forms.Application;
using System.IO;

namespace Rubberduck.UI.About
{
    public class AboutControlViewModel
    {
        private readonly IVersionCheck _version;
        private readonly IWebNavigator _web;

        public AboutControlViewModel(IVersionCheck version, IWebNavigator web)
        {
            _version = version;
            _web = web;

            UriCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteUri);
            ViewLogCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteViewLog);
        }

        public string Version => string.Format(Resources.RubberduckUI.Rubberduck_AboutBuild, _version.CurrentVersion);

        public string OperatingSystem => 
            string.Format(AboutUI.AboutWindow_OperatingSystem, Environment.OSVersion.VersionString, Environment.Is64BitOperatingSystem ? "x64" : "x86");

        public string HostProduct =>
            string.Format(AboutUI.AboutWindow_HostProduct, Application.ProductName, Environment.Is64BitProcess ? "x64" : "x86");

        public string HostVersion => string.Format(AboutUI.AboutWindow_HostVersion, Application.ProductVersion);

        public string HostExecutable => string.Format(AboutUI.AboutWindow_HostExecutable,
            Path.GetFileName(Application.ExecutablePath).ToUpper()); // .ToUpper() used to convert ExceL.EXE -> EXCEL.EXE
            
        public string AboutCopyright =>
            string.Format(AboutUI.AboutWindow_Copyright, DateTime.Now.Year);

        public CommandBase UriCommand { get; }

        public CommandBase ViewLogCommand { get; }

        private void ExecuteUri(object parameter) => _web.Navigate(((Uri)parameter));

        private void ExecuteViewLog(object parameter)
        {
            var fileTarget = (FileTarget) LogManager.Configuration.FindTargetByName("file");
                    
            var logEventInfo = new LogEventInfo { TimeStamp = DateTime.Now }; 
            var fileName = fileTarget.FileName.Render(logEventInfo);
                    
            // The /select argument will only work if the path has backslashes
            fileName = fileName.Replace("/", "\\");
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select, \"{fileName}\""));
        }
    }
}

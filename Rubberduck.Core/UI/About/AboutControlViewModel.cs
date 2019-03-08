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

        public AboutControlViewModel(IVersionCheck version)
        {
            _version = version;
        }

        public string Version => string.Format(Resources.RubberduckUI.Rubberduck_AboutBuild, _version.CurrentVersion);

        public string OperatingSystem => 
            string.Format(AboutUI.AboutWindow_OperatingSystem, Environment.OSVersion.VersionString, Environment.Is64BitOperatingSystem ? "x64" : "x86");

        public string HostProduct =>
            string.Format(AboutUI.AboutWindow_HostProduct, Application.ProductName, Environment.Is64BitProcess ? "x64" : "x86");

        public string HostVersion => string.Format(AboutUI.AboutWindow_HostVersion, Application.ProductVersion);

        public string HostExecutable => string.Format(AboutUI.AboutWindow_HostExecutable,
            Path.GetFileName(Application.ExecutablePath).ToUpper()); // .ToUpper() used to convert ExceL.EXE -> EXCEL.EXE
            
        private CommandBase _uriCommand;
        public CommandBase UriCommand
        {
            get
            {
                if (_uriCommand != null)
                {
                    return _uriCommand;
                }
                return _uriCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), uri =>
                {
                    Process.Start(new ProcessStartInfo(((Uri)uri).AbsoluteUri));
                });
            }
        }

        private CommandBase _viewLogCommand;
        public CommandBase ViewLogCommand
        {
            get
            {
                if (_viewLogCommand != null)
                {
                    return _viewLogCommand;
                }
                return _viewLogCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                {
                    var fileTarget = (FileTarget) LogManager.Configuration.FindTargetByName("file");
                    
                    var logEventInfo = new LogEventInfo { TimeStamp = DateTime.Now }; 
                    var fileName = fileTarget.FileName.Render(logEventInfo);
                    
                    // The /select argument will only work if the path has backslashes
                    fileName = fileName.Replace("/", "\\");
                    Process.Start(new ProcessStartInfo("explorer.exe", $"/select, \"{fileName}\""));
                });
            }
        }

        public string AboutCopyright =>
            string.Format(AboutUI.AboutWindow_Copyright, DateTime.Now.Year);
    }
}

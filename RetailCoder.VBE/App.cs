using System.Collections.Generic;
using System.IO;
using Infralution.Localization.Wpf;
using NLog;
using Rubberduck.Common;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Forms;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VersionCheck;
using Application = System.Windows.Forms.Application;

namespace Rubberduck
{
    public sealed class App : IDisposable
    {
        private readonly IMessageBox _messageBox;
        private readonly IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;
        private readonly IRubberduckHooks _hooks;
        private readonly IVersionCheck _version;
        private readonly CommandBase _checkVersionCommand;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        private Configuration _config;

        public App(IVBE vbe, 
            IMessageBox messageBox,
            IGeneralConfigService configService,
            IAppMenu appMenus,
            IRubberduckHooks hooks,
            IVersionCheck version,
            CommandBase checkVersionCommand)
        {
            _messageBox = messageBox;
            _configService = configService;
            _appMenus = appMenus;
            _hooks = hooks;
            _version = version;
            _checkVersionCommand = checkVersionCommand;

            _configService.SettingsChanged += _configService_SettingsChanged;
            
            UiDispatcher.Initialize();
        }

        private void _configService_SettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            _config = _configService.LoadConfiguration();
            _hooks.HookHotkeys();
            // also updates the ShortcutKey text
            _appMenus.Localize();
            UpdateLoggingLevel();

            if (e.LanguageChanged)
            {
                LoadConfig();
            }
        }

        private static void EnsureLogFolderPathExists()
        {
            try
            {
                if (!Directory.Exists(ApplicationConstants.LOG_FOLDER_PATH))
                {
                    Directory.CreateDirectory(ApplicationConstants.LOG_FOLDER_PATH);
                }
            }
            catch
            {
                //Does this need to display some sort of dialog?
            }
        }

        private static void EnsureTempPathExists()
        {
            // This is required by the parser - allow this to throw. 
            if (!Directory.Exists(ApplicationConstants.RUBBERDUCK_TEMP_PATH))
            {
                Directory.CreateDirectory(ApplicationConstants.RUBBERDUCK_TEMP_PATH);
            }
            // The parser swallows the error if deletions fail - clean up any temp files on startup
            foreach (var file in new DirectoryInfo(ApplicationConstants.RUBBERDUCK_TEMP_PATH).GetFiles())
            {
                try
                {
                        file.Delete();
                }
                catch
                {
                    // Yeah, don't care here either.
                }
            }
        }

        private void UpdateLoggingLevel()
        {
            LogLevelHelper.SetMinimumLogLevel(LogLevel.FromOrdinal(_config.UserSettings.GeneralSettings.MinimumLogLevel));
        }

        public void Startup()
        {
            EnsureLogFolderPathExists();
            EnsureTempPathExists();
            LogRubberduckSart();
            LoadConfig();
            CheckForLegacyIndenterSettings();
            _appMenus.Initialize();
            _hooks.HookHotkeys(); // need to hook hotkeys before we localize menus, to correctly display ShortcutTexts
            _appMenus.Localize();

            UpdateLoggingLevel();

            if (_config.UserSettings.GeneralSettings.CheckVersion)
            {
                _checkVersionCommand.Execute(null);
            }
        }

        public void Shutdown()
        {
            try
            {
                Debug.WriteLine("App calling Hooks.Detach.");
                _hooks.Detach();
            }
            catch
            {
                // Won't matter anymore since we're shutting everything down anyway.
            }
        }

        private void LoadConfig()
        {
            _config = _configService.LoadConfiguration();

            var currentCulture = RubberduckUI.Culture;
            try
            {
                CultureManager.UICulture = CultureInfo.GetCultureInfo(_config.UserSettings.GeneralSettings.Language.Code);
                RubberduckUI.Culture = CultureInfo.CurrentUICulture;
                InspectionsUI.Culture = CultureInfo.CurrentUICulture;
                _appMenus.Localize();
            }
            catch (CultureNotFoundException exception)
            {
                Logger.Error(exception, "Error Setting Culture for Rubberduck");
                _messageBox.Show(exception.Message, "Rubberduck", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _config.UserSettings.GeneralSettings.Language.Code = currentCulture.Name;
                _configService.SaveConfiguration(_config);
            }
        }

        private void CheckForLegacyIndenterSettings()
        {
            try
            {
                Logger.Trace("Checking for legacy Smart Indenter settings.");
                if (_config.UserSettings.GeneralSettings.SmartIndenterPrompted ||
                    !_config.UserSettings.IndenterSettings.LegacySettingsExist())
                {
                    return;
                }
                var response =
                    _messageBox.Show(RubberduckUI.SmartIndenter_LegacySettingPrompt, "Rubberduck", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (response == DialogResult.Yes)
                {
                    Logger.Trace("Attempting to load legacy Smart Indenter settings.");
                    _config.UserSettings.IndenterSettings.LoadLegacyFromRegistry();
                }
                _config.UserSettings.GeneralSettings.SmartIndenterPrompted = true;
                _configService.SaveConfiguration(_config);
            }
            catch 
            {
                //Meh.
            }
        }

        private void LogRubberduckSart()
        {
            var version = _version.CurrentVersion;
            GlobalDiagnosticsContext.Set("RubberduckVersion", version.ToString());
            var headers = new List<string>
            {
                string.Format("Rubberduck version {0} loading:", version),
                string.Format("\tOperating System: {0} {1}", Environment.OSVersion.VersionString, Environment.Is64BitOperatingSystem ? "x64" : "x86"),
                string.Format("\tHost Product: {0} {1}", Application.ProductName, Environment.Is64BitProcess ? "x64" : "x86"),
                string.Format("\tHost Version: {0}", Application.ProductVersion),
                string.Format("\tHost Executable: {0}", Path.GetFileName(Application.ExecutablePath)),
            };
            Logger.Log(LogLevel.Info, string.Join(Environment.NewLine, headers));
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= _configService_SettingsChanged;
            }

            UiDispatcher.Shutdown();

            _disposed = true;
        }
    }
}

using System.Collections.Generic;
using System.IO;
using Infralution.Localization.Wpf;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Settings;
using Rubberduck.UI.Command.MenuItems;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Resources;
using Rubberduck.Runtime;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.Utility;
using Rubberduck.VersionCheck;
using Application = System.Windows.Forms.Application;
using Rubberduck.SettingsProvider;

namespace Rubberduck
{
    public sealed class App : IDisposable
    {
        private readonly IMessageBox _messageBox;
        private readonly IConfigurationService<Configuration> _configService;
        private readonly IAppMenu _appMenus;
        private readonly IRubberduckHooks _hooks;
        private readonly IVersionCheck _version;
        private readonly CommandBase _checkVersionCommand;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        private Configuration _config;

        public App(IMessageBox messageBox,
            IConfigurationService<Configuration> configService,
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

            UiContextProvider.Initialize();
        }

        private void _configService_SettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            _config = _configService.Read();
            _hooks.HookHotkeys();
            UpdateLoggingLevel();

            if (e.LanguageChanged)
            {
                ApplyCultureConfig();
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

        /// <summary>
        /// Ensure that log level is changed to "none" after a successful
        /// run of Rubberduck for first time. By default, we ship with 
        /// log level set to Trace (0) but once it's installed and has
        /// ran without problem, it should be set to None (6)
        /// </summary>
        private void UpdateLoggingLevelOnShutdown()
        {
            if (_config.UserSettings.GeneralSettings.UserEditedLogLevel ||
                _config.UserSettings.GeneralSettings.MinimumLogLevel != LogLevel.Trace.Ordinal)
            {
                return;
            }

            _config.UserSettings.GeneralSettings.MinimumLogLevel = LogLevel.Off.Ordinal;
            _configService.Save(_config);
        }

        public void Startup()
        {
            EnsureLogFolderPathExists();
            EnsureTempPathExists();
            ApplyCultureConfig();

            LogRubberduckStart();
            UpdateLoggingLevel();
            
            CheckForLegacyIndenterSettings();
            _appMenus.Initialize();
            _hooks.HookHotkeys(); // need to hook hotkeys before we localize menus, to correctly display ShortcutTexts            
            _appMenus.Localize();

            if (_config.UserSettings.GeneralSettings.CanCheckVersion)
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

                UpdateLoggingLevelOnShutdown();
            }
            catch
            {
                // Won't matter anymore since we're shutting everything down anyway.
            }
        }

        private void ApplyCultureConfig()
        {
            _config = _configService.Read();

            var currentCulture = Resources.RubberduckUI.Culture;
            try
            {
                CultureManager.UICulture = CultureInfo.GetCultureInfo(_config.UserSettings.GeneralSettings.Language.Code);
                LocalizeResources(CultureManager.UICulture);

                _appMenus.Localize();
            }
            catch (CultureNotFoundException exception)
            {
                Logger.Error(exception, "Error Setting Culture for Rubberduck");
                // not accessing resources here, because setting resource culture literally just failed.
                _messageBox.NotifyWarn(exception.Message, "Rubberduck");
                _config.UserSettings.GeneralSettings.Language.Code = currentCulture.Name;
                _configService.Save(_config);
            }
        }

        private static void LocalizeResources(CultureInfo culture)
        {
            var localizers = AppDomain.CurrentDomain.GetAssemblies()
                .SingleOrDefault(assembly => assembly.GetName().Name == "Rubberduck.Resources")
                ?.DefinedTypes.SelectMany(type => type.DeclaredProperties.Where(prop =>
                    prop.CanWrite && prop.Name.Equals("Culture") && prop.PropertyType == typeof(CultureInfo) &&
                    (prop.SetMethod?.IsStatic ?? false)));

            if (localizers == null)
            {
                return;
            }

            var args = new object[] { culture };
            foreach (var localizer in localizers)
            {
                localizer.SetMethod.Invoke(null, args);
            }
        }

        private void CheckForLegacyIndenterSettings()
        {
            try
            {
                Logger.Trace("Checking for legacy Smart Indenter settings.");
                if (_config.UserSettings.GeneralSettings.IsSmartIndenterPrompted ||
                    !_config.UserSettings.IndenterSettings.LegacySettingsExist())
                {
                    return;
                }
                if (_messageBox.Question(Resources.RubberduckUI.SmartIndenter_LegacySettingPrompt, "Rubberduck"))
                {
                    Logger.Trace("Attempting to load legacy Smart Indenter settings.");
                    _config.UserSettings.IndenterSettings.LoadLegacyFromRegistry();
                }
                _config.UserSettings.GeneralSettings.IsSmartIndenterPrompted = true;
                _configService.Save(_config);
            }
            catch 
            {
                //Meh.
            }
        }

        public void LogRubberduckStart()
        {
            var version = _version.CurrentVersion;
            GlobalDiagnosticsContext.Set("RubberduckVersion", version.ToString());

            var headers = new List<string>
            {
                $"\r\n\tRubberduck version {version} loading:",
                $"\tOperating System: {Environment.OSVersion.VersionString} {(Environment.Is64BitOperatingSystem ? "x64" : "x86")}"
            };
            try
            {
                headers.AddRange(new []
                {
                    $"\tHost Product: {Application.ProductName} {(Environment.Is64BitProcess ? "x64" : "x86")}",
                    $"\tHost Version: {Application.ProductVersion}",
                    $"\tHost Executable: {Path.GetFileName(Application.ExecutablePath).ToUpper()}", // .ToUpper() used to convert ExceL.EXE -> EXCEL.EXE
                });
            }
            catch
            {
                headers.Add("\tHost could not be determined.");
            }

            LogLevelHelper.SetDebugInfo(string.Join(Environment.NewLine, headers));
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

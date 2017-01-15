using System.Collections.Generic;
using System.IO;
using Infralution.Localization.Wpf;
using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using System;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.UI.Command.MenuItems.CommandBars;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck
{
    public sealed class App : IDisposable
    {
        private readonly IVBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IParseCoordinator _parser;
        private readonly AutoSave.AutoSave _autoSave;
        private readonly IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;
        private readonly RubberduckCommandBar _stateBar;
        private readonly IRubberduckHooks _hooks;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        private Configuration _config;

        public App(IVBE vbe, 
            IMessageBox messageBox,
            IParseCoordinator parser,
            IGeneralConfigService configService,
            IAppMenu appMenus,
            RubberduckCommandBar stateBar,
            IRubberduckHooks hooks)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _parser = parser;
            _configService = configService;
            _autoSave = new AutoSave.AutoSave(_vbe, _configService);
            _appMenus = appMenus;
            _stateBar = stateBar;
            _hooks = hooks;

            _hooks.MessageReceived += _hooks_MessageReceived;
            _configService.SettingsChanged += _configService_SettingsChanged;
            _parser.State.StateChanged += Parser_StateChanged;
            _parser.State.StatusMessageUpdate += State_StatusMessageUpdate;

            UiDispatcher.Initialize();
        }

        private void State_StatusMessageUpdate(object sender, RubberduckStatusMessageEventArgs e)
        {
            var message = e.Message;
            if (message == ParserState.LoadingReference.ToString())
            {
                // note: ugly hack to enable Rubberduck.Parsing assembly to do this
                message = RubberduckUI.ParserState_LoadingReference;
            }

            _stateBar.SetStatusLabelCaption(message, _parser.State.ModuleExceptions.Count);
        }

        private void _hooks_MessageReceived(object sender, HookEventArgs e)
        {
            RefreshSelection();
        }

        private ParserState _lastStatus;
        private Declaration _lastSelectedDeclaration;

        private void RefreshSelection()
        {
            var pane = _vbe.ActiveCodePane;
            {
                Declaration selectedDeclaration = null;
                if (!pane.IsWrappingNullReference)
                {
                    selectedDeclaration = _parser.State.FindSelectedDeclaration(pane);
                    var refCount = selectedDeclaration == null ? 0 : selectedDeclaration.References.Count();
                    var caption = _stateBar.GetContextSelectionCaption(_vbe.ActiveCodePane, selectedDeclaration);
                    _stateBar.SetContextSelectionCaption(caption, refCount);
                }

                var currentStatus = _parser.State.Status;
                if (ShouldEvaluateCanExecute(selectedDeclaration, currentStatus))
                {
                    _appMenus.EvaluateCanExecute(_parser.State);
                    _stateBar.EvaluateCanExecute(_parser.State);
                }

                _lastStatus = currentStatus;
                _lastSelectedDeclaration = selectedDeclaration;
            }
        }

        private bool ShouldEvaluateCanExecute(Declaration selectedDeclaration, ParserState currentStatus)
        {
            return _lastStatus != currentStatus ||
                   (selectedDeclaration != null && !selectedDeclaration.Equals(_lastSelectedDeclaration)) ||
                   (selectedDeclaration == null && _lastSelectedDeclaration != null);
        }

        private void _configService_SettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            _config = _configService.LoadConfiguration();
            _hooks.HookHotkeys();
            // also updates the ShortcutKey text
            _appMenus.Localize();
            _stateBar.Localize();
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

        private void UpdateLoggingLevel()
        {
            LogLevelHelper.SetMinimumLogLevel(LogLevel.FromOrdinal(_config.UserSettings.GeneralSettings.MinimumLogLevel));
        }

        public void Startup()
        {
            EnsureLogFolderPathExists();
            LogRubberduckSart();
            LoadConfig();
            CheckForLegacyIndenterSettings();
            _appMenus.Initialize();
            _stateBar.Initialize();
            _hooks.HookHotkeys(); // need to hook hotkeys before we localize menus, to correctly display ShortcutTexts
            _appMenus.Localize();
            _stateBar.SetStatusLabelCaption(ParserState.Pending);
            _stateBar.EvaluateCanExecute(_parser.State);
            UpdateLoggingLevel();
        }

        public void Shutdown()
        {
            try
            {
                _hooks.Detach();
            }
            catch
            {
                // Won't matter anymore since we're shutting everything down anyway.
            }
        }

        private void Parser_StateChanged(object sender, EventArgs e)
        {
            Logger.Debug("App handles StateChanged ({0}), evaluating menu states...", _parser.State.Status);
            _appMenus.EvaluateCanExecute(_parser.State);
            _stateBar.EvaluateCanExecute(_parser.State);
            _stateBar.SetStatusLabelCaption(_parser.State.Status, _parser.State.ModuleExceptions.Count);
        }

        private void LoadConfig()
        {
            _config = _configService.LoadConfiguration();
            _autoSave.ConfigServiceSettingsChanged(this, EventArgs.Empty);

            var currentCulture = RubberduckUI.Culture;
            try
            {
                CultureManager.UICulture = CultureInfo.GetCultureInfo(_config.UserSettings.GeneralSettings.Language.Code);
                _appMenus.Localize();
                _stateBar.Localize();
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
            var version = GetType().Assembly.GetName().Version.ToString();
            GlobalDiagnosticsContext.Set("RubberduckVersion", version);
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

            if (_parser != null && _parser.State != null)
            {
                _parser.State.StateChanged -= Parser_StateChanged;
                _parser.State.StatusMessageUpdate -= State_StatusMessageUpdate;
            }

            if (_hooks != null)
            {
                _hooks.MessageReceived -= _hooks_MessageReceived;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= _configService_SettingsChanged;
            }

            if (_autoSave != null)
            {
                _autoSave.Dispose();
            }

            UiDispatcher.Shutdown();

            _disposed = true;
        }
    }
}

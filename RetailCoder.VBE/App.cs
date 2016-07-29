using System.IO;
using Infralution.Localization.Wpf;
using Microsoft.Vbe.Interop;
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
using System.Windows.Forms;

namespace Rubberduck
{
    public sealed class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IRubberduckParser _parser;
        private AutoSave.AutoSave _autoSave;
        private IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;
        private RubberduckCommandBar _stateBar;
        private IRubberduckHooks _hooks;
        private readonly UI.Settings.Settings _settings;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        private Configuration _config;

        public App(VBE vbe, IMessageBox messageBox,
            UI.Settings.Settings settings,
            IRubberduckParser parser,
            IGeneralConfigService configService,
            IAppMenu appMenus,
            RubberduckCommandBar stateBar,
            IRubberduckHooks hooks)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _settings = settings;
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
            _stateBar.Refresh += _stateBar_Refresh;
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

            _stateBar.SetStatusText(message);
        }

        private void _hooks_MessageReceived(object sender, HookEventArgs e)
        {
            RefreshSelection();
        }

        private ParserState _lastStatus;
        private Declaration _lastSelectedDeclaration;

        private void RefreshSelection()
        {
            var selectedDeclaration = _parser.State.FindSelectedDeclaration(_vbe.ActiveCodePane);
            _stateBar.SetSelectionText(selectedDeclaration);

            var currentStatus = _parser.State.Status;
            if (ShouldEvaluateCanExecute(selectedDeclaration, currentStatus))
            {
                _appMenus.EvaluateCanExecute(_parser.State);
            }

            _lastStatus = currentStatus;
            _lastSelectedDeclaration = selectedDeclaration;
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
            UpdateLoggingLevel();

            if (e.LanguageChanged)
            {
                LoadConfig();
            }
        }

        private void EnsureDirectoriesExist()
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
            EnsureDirectoriesExist();
            LoadConfig();
            _appMenus.Initialize();
            _hooks.HookHotkeys(); // need to hook hotkeys before we localize menus, to correctly display ShortcutTexts
            _appMenus.Localize();
            UpdateLoggingLevel();

            if (_vbe.VBProjects.Count != 0)
            {
                _parser.State.OnParseRequested(this);
            }
        }

        private void _stateBar_Refresh(object sender, EventArgs e)
        {
            // handles "refresh" button click on "Rubberduck" command bar
            _parser.State.OnParseRequested(sender);
        }

        private void Parser_StateChanged(object sender, EventArgs e)
        {
            Logger.Debug("App handles StateChanged ({0}), evaluating menu states...", _parser.State.Status);
            _appMenus.EvaluateCanExecute(_parser.State);
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
            }
            catch (CultureNotFoundException exception)
            {
                Logger.Error(exception, "Error Setting Culture for Rubberduck");
                _messageBox.Show(exception.Message, "Rubberduck", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _config.UserSettings.GeneralSettings.Language.Code = currentCulture.Name;
                _configService.SaveConfiguration(_config);
            }
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
                _parser.Dispose();
                // I won't set this to null because other components may try to release things
            }

            if (_hooks != null)
            {
                try
                {
                    _hooks.Detach();
                }
                catch {} // Won't matter anymore since we're shutting everything down anyway.

                _hooks.MessageReceived -= _hooks_MessageReceived;
                _hooks.Dispose();
                _hooks = null;
            }

            if (_settings != null)
            {
                _settings.Dispose();
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= _configService_SettingsChanged;
                _configService = null;
            }

            if (_stateBar != null)
            {
                _stateBar.Refresh -= _stateBar_Refresh;
                _stateBar.Dispose();
                _stateBar = null;
            }

            if (_autoSave != null)
            {
                _autoSave.Dispose();
                _autoSave = null;
            }

            UiDispatcher.Shutdown();

            _disposed = true;
        }
    }
}

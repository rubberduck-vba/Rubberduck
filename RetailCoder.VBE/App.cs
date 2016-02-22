using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.AutoSave;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.ParserErrors;
using Rubberduck.VBEditor.Extensions;
using Infralution.Localization.Wpf;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IRubberduckParser _parser;
        private readonly AutoSave.AutoSave _autoSave;
        private readonly IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;
        private readonly RubberduckCommandBar _stateBar;
        private readonly IIndenter _indenter;
        private readonly IRubberduckHooks _hooks;

        private readonly Logger _logger;

        private Configuration _config;

        public App(VBE vbe, IMessageBox messageBox,
            IRubberduckParser parser,
            IGeneralConfigService configService,
            IAppMenu appMenus,
            RubberduckCommandBar stateBar,
            IIndenter indenter/*,
            IRubberduckHooks hooks*/)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _parser = parser;
            _configService = configService;
            _autoSave = new AutoSave.AutoSave(_vbe, _configService);
            _appMenus = appMenus;
            _stateBar = stateBar;
            _indenter = indenter;
            //_hooks = hooks;
            _logger = LogManager.GetCurrentClassLogger();

            //_hooks.MessageReceived += hooks_MessageReceived;
            _configService.LanguageChanged += ConfigServiceLanguageChanged;
            _parser.State.StateChanged += Parser_StateChanged;
            _stateBar.Refresh += _stateBar_Refresh;

            UiDispatcher.Initialize();
        }

        private Keys _firstStepHotKey;
        private bool _isAwaitingTwoStepKey;
        private bool _skipKeyUp;

        private async void hooks_MessageReceived(object sender, HookEventArgs e)
        {
            if (sender is LowLevelKeyboardHook)
            {
                if (_skipKeyUp)
                {
                    _skipKeyUp = false;
                    return;
                }

                if (_isAwaitingTwoStepKey)
                {
                    // todo: use _firstStepHotKey and e.Key to run 2-step hotkey action
                    if (_firstStepHotKey == Keys.I && e.Key == Keys.M)
                    {
                        _indenter.IndentCurrentModule();
                    }

                    AwaitNextKey();
                    return;
                }

                var component = _vbe.ActiveCodePane.CodeModule.Parent;
                _parser.ParseComponent(component);

                AwaitNextKey();
                return;
            }

            var hotKey = sender as IHotKey;
            if (hotKey == null)
            {
                AwaitNextKey();
                return;
            }

            if (hotKey.IsTwoStepHotKey)
            {
                _firstStepHotKey = hotKey.HotKeyInfo.Keys;
                AwaitNextKey(true, hotKey.HotKeyInfo);
            }
            else
            {
                // todo: use e.Key to run 1-step hotkey action
                _firstStepHotKey = Keys.None;
                AwaitNextKey();
            }
        }

        private void AwaitNextKey(bool eatNextKey = false, HotKeyInfo info = default(HotKeyInfo))
        {
            _isAwaitingTwoStepKey = eatNextKey;
            foreach (var hook in _hooks.Hooks.OfType<ILowLevelKeyboardHook>())
            {
                hook.EatNextKey = eatNextKey;
            }

            _skipKeyUp = eatNextKey;
            if (eatNextKey)
            {
                _stateBar.SetStatusText("(" + info + ") was pressed. Waiting for second key...");
            }
            else
            {
                _stateBar.SetStatusText(_parser.State.Status.ToString());
            }
        }

        private void _stateBar_Refresh(object sender, EventArgs e)
        {
            _parser.State.OnParseRequested();
        }

        private void Parser_StateChanged(object sender, ParserStateEventArgs e)
        {
            _appMenus.EvaluateCanExecute(_parser.State);
        }

        public void Startup()
        {
            CleanReloadConfig();

            _appMenus.Initialize();
            _appMenus.Localize();

            // delay to allow the VBE to properly load. HostApplication is null until then.
            Task.Delay(1000).ContinueWith(t =>
            {
                _parser.State.AddBuiltInDeclarations(_vbe.HostApplication());
                _parser.State.OnParseRequested();
            });

            //_hooks.AddHook(new LowLevelKeyboardHook(_vbe));
            //_hooks.AddHook(new HotKey((IntPtr)_vbe.MainWindow.HWnd, "%^R", Keys.R));
            //_hooks.AddHook(new HotKey((IntPtr)_vbe.MainWindow.HWnd, "%^I", Keys.I));
            //_hooks.Attach();
        }

        private void CleanReloadConfig()
        {
            LoadConfig();
        }

        private void ConfigServiceLanguageChanged(object sender, EventArgs e)
        {
            CleanReloadConfig();
        }

        private void LoadConfig()
        {
            _logger.Debug("Loading configuration");
            _config = _configService.LoadConfiguration();

            var currentCulture = RubberduckUI.Culture;
            try
            {
                CultureManager.UICulture = CultureInfo.GetCultureInfo(_config.UserSettings.GeneralSettings.Language.Code);
                _appMenus.Localize();
            }
            catch (CultureNotFoundException exception)
            {
                _logger.Error(exception, "Error Setting Culture for RubberDuck");
                _messageBox.Show(exception.Message, "Rubberduck", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _config.UserSettings.GeneralSettings.Language.Code = currentCulture.Name;
                _configService.SaveConfiguration(_config);
            }
        }

        public void Dispose()
        {
            //_hooks.MessageReceived -= hooks_MessageReceived;
            _configService.LanguageChanged -= ConfigServiceLanguageChanged;
            _parser.State.StateChanged -= Parser_StateChanged;
            _autoSave.Dispose();

            //_hooks.Dispose();
        }
    }
}

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.ParserErrors;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck
{
    public class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IParserErrorsPresenterFactory _parserErrorsPresenterFactory;
        private readonly IRubberduckParser _parser;
        private readonly IInspectorFactory _inspectorFactory;
        private readonly IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;
        private readonly ParserStateCommandBar _stateBar;
        private readonly IIndenter _indenter;
        private readonly IRubberduckHooks _hooks;

        private readonly Logger _logger;

        private Configuration _config;

        private readonly ConcurrentDictionary<VBComponent, CancellationTokenSource> _tokenSources =
            new ConcurrentDictionary<VBComponent, CancellationTokenSource>(); 

        public App(VBE vbe, IMessageBox messageBox,
            IParserErrorsPresenterFactory parserErrorsPresenterFactory,
            IRubberduckParser parser,
            IInspectorFactory inspectorFactory, 
            IGeneralConfigService configService,
            IAppMenu appMenus,
            ParserStateCommandBar stateBar,
            IIndenter indenter,
            IRubberduckHooks hooks)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _parserErrorsPresenterFactory = parserErrorsPresenterFactory;
            _parser = parser;
            _inspectorFactory = inspectorFactory;
            _configService = configService;
            _appMenus = appMenus;
            _stateBar = stateBar;
            _indenter = indenter;
            _hooks = hooks;
            _logger = LogManager.GetCurrentClassLogger();

            _hooks.MessageReceived += hooks_MessageReceived;
            _configService.SettingsChanged += _configService_SettingsChanged;
            _parser.State.StateChanged += Parser_StateChanged;
            _stateBar.Refresh += _stateBar_Refresh;
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

                    AwaitNextKey();
                    return;
                }

                var component = _vbe.ActiveCodePane.CodeModule.Parent;
                await ParseComponentAsync(component);

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
            ParseAll();
        }

        private void Parser_StateChanged(object sender, EventArgs e)
        {
            _appMenus.EvaluateCanExecute(_parser.State);
        }

        private async Task ParseComponentAsync(VBComponent component, bool resolve = true)
        {
            var tokenSource = RenewTokenSource(component);

            var token = tokenSource.Token;
            await _parser.ParseAsync(component, token);

            if (resolve && !token.IsCancellationRequested)
            {
                using (var source = new CancellationTokenSource())
                {
                    _parser.Resolve(source.Token);
                }
            }
        }

        private CancellationTokenSource RenewTokenSource(VBComponent component)
        {
            if (_tokenSources.ContainsKey(component))
            {
                CancellationTokenSource existingTokenSource;
                _tokenSources.TryRemove(component, out existingTokenSource);
                if (existingTokenSource != null)
                {
                    existingTokenSource.Cancel();
                    existingTokenSource.Dispose();
                }
            }

            var tokenSource = new CancellationTokenSource();
            _tokenSources[component] = tokenSource;
            return tokenSource;
        }

        public void Startup()
        {
            CleanReloadConfig();

            _appMenus.Initialize();
            _appMenus.Localize();

            Task.Delay(1000).ContinueWith(t =>
            {
                _parser.State.AddBuiltInDeclarations(_vbe.HostApplication());
                ParseAll();
            });

            _hooks.AddHook(new LowLevelKeyboardHook(_vbe));
            _hooks.AddHook(new HotKey((IntPtr)_vbe.MainWindow.HWnd, "%+R", Keys.R));
            _hooks.AddHook(new HotKey((IntPtr)_vbe.MainWindow.HWnd, "%+I", Keys.I));
            _hooks.Attach();
        }

        private void ParseAll()
        {
            var components = _vbe.VBProjects.Cast<VBProject>()
                .SelectMany(project => project.VBComponents.Cast<VBComponent>());

            var result = Parallel.ForEach(components, component => { ParseComponentAsync(component, false); });

            if (result.IsCompleted)
            {
                using (var tokenSource = new CancellationTokenSource())
                {
                    _parser.Resolve(tokenSource.Token);
                }
            }
        }

        private void CleanReloadConfig()
        {
            LoadConfig();
            Setup();
        }

        private void _configService_SettingsChanged(object sender, EventArgs e)
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
                RubberduckUI.Culture = CultureInfo.GetCultureInfo(_config.UserSettings.LanguageSetting.Code);
                _appMenus.Localize();
            }
            catch (CultureNotFoundException exception)
            {
                _logger.Error(exception, "Error Setting Culture for RubberDuck");
                _messageBox.Show(exception.Message, "Rubberduck", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _config.UserSettings.LanguageSetting.Code = currentCulture.Name;
                _configService.SaveConfiguration(_config);
            }
        }

        private void Setup()
        {
            _inspectorFactory.Create();
            _parserErrorsPresenterFactory.Create();
        }

        public void Dispose()
        {
            _hooks.MessageReceived -= hooks_MessageReceived;
            _configService.SettingsChanged -= _configService_SettingsChanged;
            _parser.State.StateChanged -= Parser_StateChanged;

            _hooks.Dispose();

            if (_tokenSources.Any())
            {
                foreach (var tokenSource in _tokenSources)
                {
                    tokenSource.Value.Cancel();
                    tokenSource.Value.Dispose();
                }
            }
        }
    }
}

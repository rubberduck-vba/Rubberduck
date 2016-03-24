using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using Infralution.Localization.Wpf;
using Rubberduck.Common.Dispatch;

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

        private readonly IConnectionPoint _projectsEventsConnectionPoint;
        private readonly int _projectsEventsCookie;

        private readonly IDictionary<VBComponents, Tuple<IConnectionPoint, int>>  _componentsEventsConnectionPoints = 
            new Dictionary<VBComponents, Tuple<IConnectionPoint, int>>(); 

        public App(VBE vbe, IMessageBox messageBox,
            IRubberduckParser parser,
            IGeneralConfigService configService,
            IAppMenu appMenus,
            RubberduckCommandBar stateBar,
            IIndenter indenter/*
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

            var sink = new VBProjectsEventsSink();
            var connectionPointContainer = (IConnectionPointContainer)_vbe.VBProjects;
            var interfaceId = typeof (_dispVBProjectsEvents).GUID;
            connectionPointContainer.FindConnectionPoint(ref interfaceId, out _projectsEventsConnectionPoint);
            
            sink.ProjectAdded += sink_ProjectAdded;
            sink.ProjectRemoved += sink_ProjectRemoved;
            sink.ProjectActivated += sink_ProjectActivated;
            sink.ProjectRenamed += sink_ProjectRenamed;

            _projectsEventsConnectionPoint.Advise(sink, out _projectsEventsCookie);

            UiDispatcher.Initialize();
        }

        async void sink_ProjectRemoved(object sender, DispatcherEventArgs<VBProject> e)
        {
            Debug.WriteLine(string.Format("Project '{0}' was removed.", e.Item.Name));
            Tuple<IConnectionPoint, int> value;
            if (_componentsEventsConnectionPoints.TryGetValue(e.Item.VBComponents, out value))
            {
                value.Item1.Unadvise(value.Item2);
                _componentsEventsConnectionPoints.Remove(e.Item.VBComponents);

                _parser.State.ClearDeclarations(e.Item);
            }
        }

        async void sink_ProjectAdded(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                // forces menus to evaluate their CanExecute state:
                Parser_StateChanged(this, new ParserStateEventArgs(ParserState.Pending));
                _stateBar.SetStatusText();
                return;
            }

            Debug.WriteLine(string.Format("Project '{0}' was added.", e.Item.Name));
            var connectionPointContainer = (IConnectionPointContainer)e.Item.VBComponents;
            var interfaceId = typeof(_dispVBComponentsEvents).GUID;
            
            IConnectionPoint connectionPoint;
            connectionPointContainer.FindConnectionPoint(ref interfaceId, out connectionPoint);

            var sink = new VBComponentsEventsSink();
            sink.ComponentActivated += sink_ComponentActivated;
            sink.ComponentAdded += sink_ComponentAdded;
            sink.ComponentReloaded += sink_ComponentReloaded;
            sink.ComponentRemoved += sink_ComponentRemoved;
            sink.ComponentRenamed += sink_ComponentRenamed;
            sink.ComponentSelected += sink_ComponentSelected;

            int cookie;
            connectionPoint.Advise(sink, out cookie);

            _componentsEventsConnectionPoints.Add(e.Item.VBComponents, Tuple.Create(connectionPoint, cookie));
            _parser.State.OnParseRequested(sender);
        }

        async void sink_ComponentSelected(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Component '{0}' was selected.", e.Item.Name));
            // do something?
        }

        async void sink_ComponentRenamed(object sender, DispatcherRenamedEventArgs<VBComponent> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Component '{0}' was renamed.", e.Item.Name));

            _parser.State.OnParseRequested(sender, e.Item);
        }

        async void sink_ComponentRemoved(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Component '{0}' was removed.", e.Item.Name));
            _parser.State.ClearDeclarations(e.Item);
        }

        async void sink_ComponentReloaded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Component '{0}' was reloaded.", e.Item.Name));
            _parser.State.OnParseRequested(sender, e.Item);
        }

        async void sink_ComponentAdded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Component '{0}' was added.", e.Item.Name));
            _parser.State.OnParseRequested(sender, e.Item);
        }

        async void sink_ComponentActivated(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Component '{0}' was activated.", e.Item.Name));
            // do something?
        }

        async void sink_ProjectRenamed(object sender, DispatcherRenamedEventArgs<VBProject> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Project '{0}' was renamed.", e.Item.Name));
            _parser.State.ClearDeclarations(e.Item);
            _parser.State.OnParseRequested(sender);
        }

        async void sink_ProjectActivated(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            Debug.WriteLine(string.Format("Project '{0}' was activated.", e.Item.Name));
            // do something?
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
            _parser.State.OnParseRequested(sender);
        }

        private void Parser_StateChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("App handles StateChanged ({0}), evaluating menu states...", _parser.State.Status);
            _appMenus.EvaluateCanExecute(_parser.State);
        }

        public void Startup()
        {
            CleanReloadConfig();

            _appMenus.Initialize();
            _appMenus.Localize();

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
            _configService.LanguageChanged -= ConfigServiceLanguageChanged;
            _parser.State.StateChanged -= Parser_StateChanged;
            _autoSave.Dispose();

            _projectsEventsConnectionPoint.Unadvise(_projectsEventsCookie);
            foreach (var item in _componentsEventsConnectionPoints)
            {
                item.Value.Item1.Unadvise(item.Value.Item2);
            }

            //_hooks.MessageReceived -= hooks_MessageReceived;
            //_hooks.Dispose();
        }
    }
}

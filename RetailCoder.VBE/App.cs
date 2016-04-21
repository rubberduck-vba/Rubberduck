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
using Rubberduck.Common.Hotkeys;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

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

        private readonly IDictionary<VBProjectsEventsSink, Tuple<IConnectionPoint, int>>  _componentsEventsConnectionPoints = 
            new Dictionary<VBProjectsEventsSink, Tuple<IConnectionPoint, int>>();
        private readonly IDictionary<VBProjectsEventsSink, Tuple<IConnectionPoint, int>> _referencesEventsConnectionPoints =
            new Dictionary<VBProjectsEventsSink, Tuple<IConnectionPoint, int>>();

        private readonly IDictionary<Type, Action> _hookActions;

        public App(VBE vbe, IMessageBox messageBox,
            IRubberduckParser parser,
            IGeneralConfigService configService,
            IAppMenu appMenus,
            RubberduckCommandBar stateBar,
            IIndenter indenter,
            IRubberduckHooks hooks)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _parser = parser;
            _configService = configService;
            _autoSave = new AutoSave.AutoSave(_vbe, _configService);
            _appMenus = appMenus;
            _stateBar = stateBar;
            _indenter = indenter;
            _hooks = hooks;
            _logger = LogManager.GetCurrentClassLogger();

            _hooks.MessageReceived += _hooks_MessageReceived;
            _configService.SettingsChanged += _configService_SettingsChanged;
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

            _hookActions = new Dictionary<Type, Action>
            {
                { typeof(MouseHook), HandleMouseMessage },
                { typeof(KeyboardHook), HandleKeyboardMessage },
            };
            
            
            UiDispatcher.Initialize();
        }

        private void _hooks_MessageReceived(object sender, HookEventArgs e)
        {
            var hookType = sender.GetType();
            Action action;
            if (_hookActions.TryGetValue(hookType, out action))
            {
                action.Invoke();
            }
        }

        private void HandleMouseMessage()
        {
            RefreshSelection();
        }

        private void HandleKeyboardMessage()
        {
            RefreshSelection();
        }

        private void RefreshSelection()
        {
            _stateBar.SetSelectionText(_parser.State.FindSelectedDeclaration(_vbe.ActiveCodePane));
            _appMenus.EvaluateCanExecute(_parser.State);
        }

        private void _configService_SettingsChanged(object sender, EventArgs e)
        {
            // also updates the ShortcutKey text
            _appMenus.Localize();
            _hooks.HookHotkeys();
        }

        public void Startup()
        {
            CleanReloadConfig();

            foreach (var project in _vbe.VBProjects.Cast<VBProject>())
            {
                _parser.State.AddProject(project);
            }

            _appMenus.Initialize();
            _appMenus.Localize();

            Task.Delay(1000).ContinueWith(t =>
            {
                // run this on UI thread
                UiDispatcher.Invoke(() =>
                {
                    _parser.State.OnParseRequested(this);
                });
            }).ConfigureAwait(false);

            _hooks.HookHotkeys();
        }

        #region sink handlers. todo: move to another class
        async void sink_ProjectRemoved(object sender, DispatcherEventArgs<VBProject> e)
        {
            var sink = (VBProjectsEventsSink)sender;
            _componentsEventSinks.Remove(sink);
            _referencesEventsSinks.Remove(sink);
            _parser.State.RemoveProject(e.Item);

            Debug.WriteLine(string.Format("Project '{0}' was removed.", e.Item.Name));
            Tuple<IConnectionPoint, int> componentsTuple;
            if (_componentsEventsConnectionPoints.TryGetValue(sink, out componentsTuple))
            {
                componentsTuple.Item1.Unadvise(componentsTuple.Item2);
                _componentsEventsConnectionPoints.Remove(sink);
            }

            Tuple<IConnectionPoint, int> referencesTuple;
            if (_referencesEventsConnectionPoints.TryGetValue(sink, out referencesTuple))
            {
                referencesTuple.Item1.Unadvise(referencesTuple.Item2);
                _referencesEventsConnectionPoints.Remove(sink);
            }

            _parser.State.ClearDeclarations(e.Item);
        }

        private readonly IDictionary<VBProjectsEventsSink, VBComponentsEventsSink> _componentsEventSinks = 
            new Dictionary<VBProjectsEventsSink, VBComponentsEventsSink>();

        private readonly IDictionary<VBProjectsEventsSink, ReferencesEventsSink> _referencesEventsSinks = 
            new Dictionary<VBProjectsEventsSink, ReferencesEventsSink>();

        async void sink_ProjectAdded(object sender, DispatcherEventArgs<VBProject> e)
        {
            var sink = (VBProjectsEventsSink)sender;
            RegisterComponentsEventSink(e, sink);
            _parser.State.AddProject(e.Item);

            if (!_parser.State.AllDeclarations.Any())
            {
                // forces menus to evaluate their CanExecute state:
                Parser_StateChanged(this, new ParserStateEventArgs(ParserState.Pending));
                _stateBar.SetStatusText();
                return;
            }

            Debug.WriteLine(string.Format("Project '{0}' was added.", e.Item.Name));
            _parser.State.OnParseRequested(sender);
        }

        private void RegisterComponentsEventSink(DispatcherEventArgs<VBProject> e, VBProjectsEventsSink sink)
        {
            var connectionPointContainer = (IConnectionPointContainer) e.Item.VBComponents;
            var interfaceId = typeof (_dispVBComponentsEvents).GUID;

            IConnectionPoint connectionPoint;
            connectionPointContainer.FindConnectionPoint(ref interfaceId, out connectionPoint);

            var componentsSink = new VBComponentsEventsSink();
            componentsSink.ComponentActivated += sink_ComponentActivated;
            componentsSink.ComponentAdded += sink_ComponentAdded;
            componentsSink.ComponentReloaded += sink_ComponentReloaded;
            componentsSink.ComponentRemoved += sink_ComponentRemoved;
            componentsSink.ComponentRenamed += sink_ComponentRenamed;
            componentsSink.ComponentSelected += sink_ComponentSelected;
            _componentsEventSinks.Add(sink, componentsSink);

            int cookie;
            connectionPoint.Advise(componentsSink, out cookie);

            _componentsEventsConnectionPoints.Add(sink, Tuple.Create(connectionPoint, cookie));
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

            Debug.WriteLine("Project '{0}' (ID {1}) was renamed to '{2}'.", e.OldName, e.Item.HelpFile, e.Item.Name);
            _parser.State.RemoveProject(e.Item.HelpFile);
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
        #endregion

        private void _stateBar_Refresh(object sender, EventArgs e)
        {
            // handles "refresh" button click on "Rubberduck" command bar
            _parser.State.OnParseRequested(sender);
        }

        private void Parser_StateChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("App handles StateChanged ({0}), evaluating menu states...", _parser.State.Status);
            _appMenus.EvaluateCanExecute(_parser.State);
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
                _logger.Error(exception, "Error Setting Culture for Rubberduck");
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

            _hooks.Dispose();
        }
    }
}

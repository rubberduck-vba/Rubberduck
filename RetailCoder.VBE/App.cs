using Infralution.Localization.Wpf;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.Common;
using Rubberduck.Common.Dispatch;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI;
using Rubberduck.UI.Command.MenuItems;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.UI.SourceControl;

namespace Rubberduck
{
    public sealed class App : IDisposable
    {
        private const string FILE_TARGET_NAME = "file";
        private readonly VBE _vbe;
        private readonly IMessageBox _messageBox;
        private IRubberduckParser _parser;
        private AutoSave.AutoSave _autoSave;
        private IGeneralConfigService _configService;
        private readonly IAppMenu _appMenus;
        private RubberduckCommandBar _stateBar;
        private IRubberduckHooks _hooks;
        private bool _handleSinkEvents = true;
        private readonly BranchesViewViewModel _branchesVM;
        private readonly SourceControlViewViewModel _panelVM;

        private readonly Logger _logger;

        private VBProjectsEventsSink _sink;
        private Configuration _config;

        private readonly IConnectionPoint _projectsEventsConnectionPoint;
        private readonly int _projectsEventsCookie;

        private readonly IDictionary<string, Tuple<IConnectionPoint, int>> _componentsEventsConnectionPoints =
            new Dictionary<string, Tuple<IConnectionPoint, int>>();
        private readonly IDictionary<string, Tuple<IConnectionPoint, int>> _referencesEventsConnectionPoints =
            new Dictionary<string, Tuple<IConnectionPoint, int>>();

        public App(VBE vbe, IMessageBox messageBox,
            IRubberduckParser parser,
            IGeneralConfigService configService,
            IAppMenu appMenus,
            RubberduckCommandBar stateBar,
            IRubberduckHooks hooks,
            SourceControlDockablePresenter sourceControlPresenter)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _parser = parser;
            _configService = configService;
            _autoSave = new AutoSave.AutoSave(_vbe, _configService);
            _appMenus = appMenus;
            _stateBar = stateBar;
            _hooks = hooks;
            _logger = LogManager.GetCurrentClassLogger();

            var sourceControlPanel = (SourceControlPanel) sourceControlPresenter.Window();
            _panelVM = (SourceControlViewViewModel) sourceControlPanel.ViewModel;
            _branchesVM = (BranchesViewViewModel) _panelVM.TabItems.Single(t => t.ViewModel.Tab == SourceControlTab.Branches).ViewModel;

            _panelVM.OpenRepoStarted += DisableSinkEventHandlers;
            _panelVM.OpenRepoCompleted += EnableSinkEventHandlersAndUpdateCache;

            _branchesVM.LoadingComponentsStarted += DisableSinkEventHandlers;
            _branchesVM.LoadingComponentsCompleted += EnableSinkEventHandlersAndUpdateCache;

            _hooks.MessageReceived += _hooks_MessageReceived;
            _configService.SettingsChanged += _configService_SettingsChanged;
            _configService.LanguageChanged += ConfigServiceLanguageChanged;
            _parser.State.StateChanged += Parser_StateChanged;
            _parser.State.StatusMessageUpdate += State_StatusMessageUpdate;
            _stateBar.Refresh += _stateBar_Refresh;

            _sink = new VBProjectsEventsSink();
            var connectionPointContainer = (IConnectionPointContainer)_vbe.VBProjects;
            var interfaceId = typeof(_dispVBProjectsEvents).GUID;
            connectionPointContainer.FindConnectionPoint(ref interfaceId, out _projectsEventsConnectionPoint);

            _sink.ProjectAdded += sink_ProjectAdded;
            _sink.ProjectRemoved += sink_ProjectRemoved;
            _sink.ProjectActivated += sink_ProjectActivated;
            _sink.ProjectRenamed += sink_ProjectRenamed;

            _projectsEventsConnectionPoint.Advise(_sink, out _projectsEventsCookie);
            UiDispatcher.Initialize();
        }

        private void EnableSinkEventHandlersAndUpdateCache(object sender, EventArgs e)
        {
            _handleSinkEvents = true;

            // update cache
            _parser.State.RemoveProject(_vbe.ActiveVBProject.HelpFile);
            _parser.State.AddProject(_vbe.ActiveVBProject);

            _parser.State.OnParseRequested(this);
        }

        private void DisableSinkEventHandlers(object sender, EventArgs e)
        {
            _handleSinkEvents = false;
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

        private void _configService_SettingsChanged(object sender, EventArgs e)
        {
            _config = _configService.LoadConfiguration();
            _hooks.HookHotkeys();
            // also updates the ShortcutKey text
            _appMenus.Localize();
            UpdateLoggingLevel();
        }

        private void UpdateLoggingLevel()
        {
            LogLevelHelper.SetMinimumLogLevel(LogLevel.FromOrdinal(_config.UserSettings.GeneralSettings.MinimumLogLevel));
        }

        public void Startup()
        {
            CleanReloadConfig();
            _appMenus.Initialize();
            _hooks.HookHotkeys(); // need to hook hotkeys before we localize menus, to correctly display ShortcutTexts
            _appMenus.Localize();
            Task.Delay(1000).ContinueWith(t => UiDispatcher.Invoke(() => _parser.State.OnParseRequested(this)));
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

        #region sink handlers. todo: move to another class
        async void sink_ProjectRemoved(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!_handleSinkEvents) { return; }

            if (e.Item.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                _logger.Debug("Locked project '{0}' was removed.", e.Item.Name);
                return;
            }

            var projectId = e.Item.HelpFile;
            Debug.Assert(projectId != null);

            _componentsEventsSinks.Remove(projectId);
            _referencesEventsSinks.Remove(projectId);
            _parser.State.RemoveProject(e.Item);

            _logger.Debug("Project '{0}' was removed.", e.Item.Name);
            Tuple<IConnectionPoint, int> componentsTuple;
            if (_componentsEventsConnectionPoints.TryGetValue(projectId, out componentsTuple))
            {
                componentsTuple.Item1.Unadvise(componentsTuple.Item2);
                _componentsEventsConnectionPoints.Remove(projectId);
            }

            Tuple<IConnectionPoint, int> referencesTuple;
            if (_referencesEventsConnectionPoints.TryGetValue(projectId, out referencesTuple))
            {
                referencesTuple.Item1.Unadvise(referencesTuple.Item2);
                _referencesEventsConnectionPoints.Remove(projectId);
            }
        }

        private readonly IDictionary<string, VBComponentsEventsSink> _componentsEventsSinks =
            new Dictionary<string, VBComponentsEventsSink>();

        private readonly IDictionary<string, ReferencesEventsSink> _referencesEventsSinks =
            new Dictionary<string, ReferencesEventsSink>();

        async void sink_ProjectAdded(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!_handleSinkEvents) { return; }

            _logger.Debug("Project '{0}' was added.", e.Item.Name);
            if (e.Item.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                _logger.Debug("Project is protected and will not be added to parser state.");
                return;
            }

            _parser.State.AddProject(e.Item); // note side-effect: assigns ProjectId/HelpFile
            var projectId = e.Item.HelpFile;
            RegisterComponentsEventSink(e.Item.VBComponents, projectId);

            if (!_parser.State.AllDeclarations.Any())
            {
                // forces menus to evaluate their CanExecute state:
                Parser_StateChanged(this, new ParserStateEventArgs(ParserState.Pending));
                _stateBar.SetStatusText();
                return;
            }

            _parser.State.OnParseRequested(sender);
        }

        private void RegisterComponentsEventSink(VBComponents components, string projectId)
        {
            if (_componentsEventsSinks.ContainsKey(projectId))
            {
                // already registered - this is caused by the initial load+rename of a project in the VBE
                _logger.Debug("Components sink already registered.");
                return;
            }

            var connectionPointContainer = (IConnectionPointContainer)components;
            var interfaceId = typeof(_dispVBComponentsEvents).GUID;

            IConnectionPoint connectionPoint;
            connectionPointContainer.FindConnectionPoint(ref interfaceId, out connectionPoint);

            var componentsSink = new VBComponentsEventsSink();
            componentsSink.ComponentActivated += sink_ComponentActivated;
            componentsSink.ComponentAdded += sink_ComponentAdded;
            componentsSink.ComponentReloaded += sink_ComponentReloaded;
            componentsSink.ComponentRemoved += sink_ComponentRemoved;
            componentsSink.ComponentRenamed += sink_ComponentRenamed;
            componentsSink.ComponentSelected += sink_ComponentSelected;
            _componentsEventsSinks.Add(projectId, componentsSink);

            int cookie;
            connectionPoint.Advise(componentsSink, out cookie);

            _componentsEventsConnectionPoints.Add(projectId, Tuple.Create(connectionPoint, cookie));
            _logger.Debug("Components sink registered and advising.");
        }

        async void sink_ComponentSelected(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _logger.Debug("Component '{0}' was selected.", e.Item.Name);
            // do something?
        }

        async void sink_ComponentRenamed(object sender, DispatcherRenamedEventArgs<VBComponent> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _logger.Debug("Component '{0}' was renamed to '{1}'.", e.OldName, e.Item.Name);

            _parser.State.RemoveRenamedComponent(e.Item, e.OldName);
        }

        async void sink_ComponentRemoved(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _panelVM.RemoveComponent(e.Item);

            _logger.Debug("Component '{0}' was removed.", e.Item.Name);
            _parser.State.ClearStateCache(e.Item, true);
        }

        async void sink_ComponentReloaded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _logger.Debug("Component '{0}' was reloaded.", e.Item.Name);
            _parser.State.OnParseRequested(sender, e.Item);
        }

        async void sink_ComponentAdded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _panelVM.AddComponent(e.Item);

            _logger.Debug("Component '{0}' was added.", e.Item.Name);
            _parser.State.OnParseRequested(sender, e.Item);
        }

        async void sink_ComponentActivated(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _logger.Debug("Component '{0}' was activated.", e.Item.Name);
            // do something?
        }

        async void sink_ProjectRenamed(object sender, DispatcherRenamedEventArgs<VBProject> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _logger.Debug("Project '{0}' (ID {1}) was renamed to '{2}'.", e.OldName, e.Item.HelpFile, e.Item.Name);

            _parser.State.RemoveProject(e.Item.HelpFile);
            _parser.State.AddProject(e.Item);

            _parser.State.OnParseRequested(sender);
        }

        async void sink_ProjectActivated(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!_handleSinkEvents) { return; }

            if (!_parser.State.AllDeclarations.Any())
            {
                return;
            }

            _logger.Debug("Project '{0}' was activated.", e.Item.Name);
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
            _logger.Debug("App handles StateChanged ({0}), evaluating menu states...", _parser.State.Status);
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

        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            if (_panelVM != null)
            {
                _panelVM.OpenRepoStarted -= DisableSinkEventHandlers;
                _panelVM.OpenRepoCompleted -= EnableSinkEventHandlersAndUpdateCache;
            }

            if (_branchesVM != null)
            {
                _branchesVM.LoadingComponentsStarted -= DisableSinkEventHandlers;
                _branchesVM.LoadingComponentsCompleted -= EnableSinkEventHandlersAndUpdateCache;
            }

            _handleSinkEvents = false;

            if (_parser != null && _parser.State != null)
            {
                _parser.State.StateChanged -= Parser_StateChanged;
                _parser.State.StatusMessageUpdate -= State_StatusMessageUpdate;
                _parser.Dispose();
                // I won't set this to null because other components may try to release things
            }

            if (_hooks != null)
            {
                _hooks.MessageReceived -= _hooks_MessageReceived;
                _hooks.Dispose();
                _hooks = null;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= _configService_SettingsChanged;
                _configService.LanguageChanged -= ConfigServiceLanguageChanged;
                _configService = null;
            }

            if (_stateBar != null)
            {
                _stateBar.Refresh -= _stateBar_Refresh;
                _stateBar.Dispose();
                _stateBar = null;
            }

            if (_sink != null)
            {
                _sink.ProjectAdded -= sink_ProjectAdded;
                _sink.ProjectRemoved -= sink_ProjectRemoved;
                _sink.ProjectActivated -= sink_ProjectActivated;
                _sink.ProjectRenamed -= sink_ProjectRenamed;
                _sink = null;
            }

            foreach (var item in _componentsEventsSinks)
            {
                item.Value.ComponentActivated -= sink_ComponentActivated;
                item.Value.ComponentAdded -= sink_ComponentAdded;
                item.Value.ComponentReloaded -= sink_ComponentReloaded;
                item.Value.ComponentRemoved -= sink_ComponentRemoved;
                item.Value.ComponentRenamed -= sink_ComponentRenamed;
                item.Value.ComponentSelected -= sink_ComponentSelected;
            }

            if (_autoSave != null)
            {
                _autoSave.Dispose();
                _autoSave = null;
            }

            _projectsEventsConnectionPoint.Unadvise(_projectsEventsCookie);
            foreach (var item in _componentsEventsConnectionPoints)
            {
                item.Value.Item1.Unadvise(item.Value.Item2);
            }
            foreach (var item in _referencesEventsConnectionPoints)
            {
                item.Value.Item1.Unadvise(item.Value.Item2);
            }

            UiDispatcher.Shutdown();

            _disposed = true;
        }
    }
}

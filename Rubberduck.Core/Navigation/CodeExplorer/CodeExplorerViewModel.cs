using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using NLog;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.Resources;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Windows;
using System.Windows.Input;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Templates;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
// ReSharper disable ExplicitCallerInfoArgument

namespace Rubberduck.Navigation.CodeExplorer
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public sealed class CodeExplorerViewModel : ViewModelBase
    {
        private readonly FolderHelper _folderHelper;
        private readonly RubberduckParserState _state;
        private readonly IConfigProvider<WindowSettings> _windowSettingsProvider;
        // ReSharper disable once NotAccessedField.Local - YGNI pending a redesign of font sizes.
        private readonly GeneralSettings _generalSettings;
        private readonly WindowSettings _windowSettings;
        private readonly IUiDispatcher _uiDispatcher;
        private readonly IVBE _vbe;
        private readonly ITemplateProvider _templateProvider;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public CodeExplorerViewModel(
            FolderHelper folderHelper,
            RubberduckParserState state,
            RemoveCommand removeCommand,
            IConfigProvider<GeneralSettings> generalSettingsProvider, 
            IConfigProvider<WindowSettings> windowSettingsProvider, 
            IUiDispatcher uiDispatcher,
            IVBE vbe,
            ITemplateProvider templateProvider,
            ICodeExplorerSyncProvider syncProvider)
        {
            _folderHelper = folderHelper;
            _state = state;
            _state.StateChanged += HandleStateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;
            _windowSettingsProvider = windowSettingsProvider;
            _uiDispatcher = uiDispatcher;
            _vbe = vbe;
            _templateProvider = templateProvider;

            if (generalSettingsProvider != null)
            {
                _generalSettings = generalSettingsProvider.Create();
            }

            if (windowSettingsProvider != null)
            {
                _windowSettings = windowSettingsProvider.Create();
            }
            CollapseAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCollapseNodes);
            ExpandAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteExpandNodes);

            _externalRemoveCommand = removeCommand;
            if (_externalRemoveCommand != null)
            {
                RemoveCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveCommand, _externalRemoveCommand.CanExecute);
            }

            SetNameSortCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                if ((bool)param)
                {
                    SortByName = (bool)param;
                    SortByCodeOrder = !(bool)param;
                }
            }, param => !SortByName);

            SetCodeOrderSortCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                if ((bool)param)
                {
                    SortByCodeOrder = (bool)param;
                    SortByName = !(bool)param;
                }
            }, param => !SortByCodeOrder);

            ClearFilterTextCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteClearSearchCommand);

            SyncCodePaneCommand = syncProvider.GetSyncCommand(this);
            // Force a call to EvaluateCanExecute
            OnPropertyChanged("SyncCodePaneCommand");
        }

        public ObservableCollection<Template> BuiltInTemplates =>
            new ObservableCollection<Template>(_templateProvider.GetTemplates().Where(t => !t.IsUserDefined)
                .OrderBy(t => t.Name));

        public ObservableCollection<Template> UserDefinedTemplates =>
            new ObservableCollection<Template>(_templateProvider.GetTemplates().Where(t => t.IsUserDefined)
                .OrderBy(t => t.Name));

        private CodeExplorerItemViewModel _selectedItem;
        public CodeExplorerItemViewModel SelectedItem
        {
            get => _selectedItem;
            set
            {
                _selectedItem = value;
                ExpandToNode(value);

                OnPropertyChanged();

                OnPropertyChanged("CanExecuteIndenterCommand");
                OnPropertyChanged("CanExecuteRenameCommand");
                OnPropertyChanged("CanExecuteFindAllReferencesCommand");
                OnPropertyChanged("ExportVisibility");
                OnPropertyChanged("ExportAllVisibility");
                OnPropertyChanged("PanelTitle");
                OnPropertyChanged("Description");
            }
        }

        public bool AnyTemplatesCanExecute =>
            BuiltInTemplates.Concat(UserDefinedTemplates)
                .Any(template => AddTemplateCommand.CanExecuteForNode(SelectedItem));
    
        public bool SortByName
        {
            get => _windowSettings.CodeExplorer_SortByName;
            set
            {
                if (_windowSettings.CodeExplorer_SortByName == value)
                {
                    return;
                }

                _windowSettings.CodeExplorer_SortByName = value;
                _windowSettings.CodeExplorer_SortByCodeOrder = !value;
                _windowSettingsProvider.Save(_windowSettings);
                OnPropertyChanged();
                OnPropertyChanged("SortByCodeOrder");

                ReorderChildNodes(Projects);
            }
        }

        public bool SortByCodeOrder
        {
            get => _windowSettings.CodeExplorer_SortByCodeOrder;
            set
            {
                if (_windowSettings.CodeExplorer_SortByCodeOrder == value)
                {
                    return;
                }

                _windowSettings.CodeExplorer_SortByCodeOrder = value;
                _windowSettings.CodeExplorer_SortByName = !value;
                _windowSettingsProvider.Save(_windowSettings);
                OnPropertyChanged();
                OnPropertyChanged("SortByName");

                ReorderChildNodes(Projects);
            }
        }

        private void ExecuteClearSearchCommand(object parameter)
        {
            if (!string.IsNullOrEmpty(FilterText))
            {
                FilterText = string.Empty;
            }
        }

        public CommandBase SetNameSortCommand { get; }

        public CommandBase SetCodeOrderSortCommand { get; }

        public bool GroupByType
        {
            get => _windowSettings.CodeExplorer_GroupByType;
            set
            {
                if (_windowSettings.CodeExplorer_GroupByType != value)
                {
                    _windowSettings.CodeExplorer_GroupByType = value;
                    _windowSettingsProvider.Save(_windowSettings);

                    OnPropertyChanged();

                    ReorderChildNodes(Projects);
                }
            }
        }

        private bool _canSearch;

        public bool CanSearch
        {
            get => _canSearch;
            set
            {
                _canSearch = value;
                OnPropertyChanged();
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                _isBusy = value;
                OnPropertyChanged();
                // If the window is "busy" then hide the Refresh message
                OnPropertyChanged("EmptyUIRefreshMessageVisibility");
            }
        }

        public string PanelTitle
        {
            get
            {
                if (SelectedItem == null)
                {
                    return string.Empty;
                }

                if (SelectedItem is CodeExplorerReferenceViewModel reference)
                {
                    return reference.Reference.Description;
                }

                var declaration = SelectedItem.Declaration;
                
                var nameWithDeclarationType = declaration.IdentifierName +
                                              $" - ({RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, CultureInfo.CurrentUICulture)})";

                if (string.IsNullOrEmpty(declaration.AsTypeName))
                {
                    return nameWithDeclarationType;
                }

                var typeName = declaration.HasTypeHint
                    ? SymbolList.TypeHintToTypeName[declaration.TypeHint]
                    : declaration.AsTypeName;

                return nameWithDeclarationType + ": " + typeName;
            }
        }

        public string Description
        {
            get
            {
                if (SelectedItem is CodeExplorerCustomFolderViewModel folder)
                {
                    return folder.FolderAttribute ?? string.Empty;
                }

                if (SelectedItem is CodeExplorerReferenceViewModel reference)
                {
                    return reference.Reference?.FullPath ?? string.Empty;
                }
                
                return SelectedItem?.Declaration?.DescriptionString ?? string.Empty;
            }
        }

        public bool CanExecuteIndenterCommand => IndenterCommand?.CanExecute(SelectedItem) ?? false;
        public bool CanExecuteRenameCommand => RenameCommand?.CanExecute(SelectedItem) ?? false;
        public bool CanExecuteFindAllReferencesCommand => FindAllReferencesCommand?.CanExecute(SelectedItem) ?? false;

        private ObservableCollection<CodeExplorerItemViewModel> _projects;
        public ObservableCollection<CodeExplorerItemViewModel> Projects
        {
            get => _projects;
            set
            {
                _projects = ForceProjectsRefresh(value);

                OnPropertyChanged();
                // Once a Project has been set, show the TreeView
                OnPropertyChanged("TreeViewVisibility");
                OnPropertyChanged("CanSearch");
            }
        }

        private ObservableCollection<CodeExplorerItemViewModel> ForceProjectsRefresh(ObservableCollection<CodeExplorerItemViewModel> projects)
        {
            ReorderChildNodes(projects);
            CanSearch = projects.Any();

            return new ObservableCollection<CodeExplorerItemViewModel>(projects.OrderBy(o => o.NameWithSignature));
        }

        private void HandleStateChanged(object sender, ParserStateEventArgs e)
        {
            if (Projects == null)
            {
                Projects = new ObservableCollection<CodeExplorerItemViewModel>();
            }

            if (e.State == ParserState.Ready)
            {
                // Finished up resolving references, so we can now update the reference nodes.
                var referenceFolders = Projects.SelectMany(node => node.Items.OfType<CodeExplorerReferenceFolderViewModel>());
                foreach (var library in referenceFolders.SelectMany(folder => folder.Items).OfType<CodeExplorerReferenceViewModel>())
                {
                    var reference = library.Reference;
                    if (reference == null)
                    {
                        continue;
                    }

                    reference.IsUsed = reference.IsBuiltIn ||
                                       _state.DeclarationFinder.IsReferenceUsedInProject(
                                           library.Parent?.Declaration as ProjectDeclaration,
                                           reference.ToReferenceInfo());
                    library.IsDimmed = !reference.IsUsed;
                }

                return;
            }

            IsBusy = _state.Status != ParserState.Pending && _state.Status <= ParserState.ResolvedDeclarations;

            if (e.State != ParserState.ResolvedDeclarations)
            {
                return;
            }

            var userDeclarations = _state.DeclarationFinder.AllUserDeclarations
                .GroupBy(declaration => declaration.ProjectId)
                .Where(grouping => grouping.Any(declaration => declaration.DeclarationType == DeclarationType.Project))
                .ToList();

            if (userDeclarations.Any(
                    grouping => grouping.All(declaration => declaration.DeclarationType != DeclarationType.Project)))
            {
                return;
            }

            var newProjects = userDeclarations.Select(grouping =>
                new CodeExplorerProjectViewModel(_folderHelper,
                    grouping.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Project),
                    grouping,
                    _vbe,
                    true)).ToList();

            UpdateNodes(Projects, newProjects);
            
            Projects = new ObservableCollection<CodeExplorerItemViewModel>(newProjects);

            FilterByName(Projects, _filterText);
        }

        private void UpdateNodes(IEnumerable<CodeExplorerItemViewModel> oldList, IEnumerable<CodeExplorerItemViewModel> newList)
        {
            var existingNodes = FlattenNodeList(oldList);
            var newNodes = FlattenNodeList(newList);

            foreach (var item in newNodes)
            {
                var projectId = item.Declaration.ProjectId;
                CodeExplorerItemViewModel matchingNode;

                switch (item)
                {
                    case CodeExplorerProjectViewModel _:
                        matchingNode = existingNodes.FirstOrDefault(node => node.Declaration.ProjectId.Equals(projectId));
                        break;
                    case CodeExplorerCustomFolderViewModel folder:
                        matchingNode = existingNodes.FirstOrDefault(node => node.Declaration.ProjectId.Equals(projectId) && node.Name == folder.Name);
                        break;
                    case CodeExplorerReferenceViewModel reference:
                    {
                        var info = reference.Reference.ToReferenceInfo();
                        matchingNode = existingNodes.OfType<CodeExplorerReferenceViewModel>().FirstOrDefault(node =>
                            node.Declaration.ProjectId.Equals(projectId) && node.Reference.Matches(info));
                        break;
                    }
                    default:
                        matchingNode = existingNodes.FirstOrDefault(node =>
                            node.Declaration.ProjectId.Equals(projectId) &&
                            item.Declaration.QualifiedName.Equals(node.Declaration.QualifiedName));
                        break;
                }

                if (matchingNode == null)
                {
                    continue;
                }

                item.IsExpanded = matchingNode.IsExpanded;
                item.IsSelected = matchingNode.IsSelected;

                if (!_unfilteredExpandedState.ContainsKey(matchingNode))
                {
                    continue;
                }

                var unfilteredState = _unfilteredExpandedState[matchingNode];
                _unfilteredExpandedState.Remove(matchingNode);
                _unfilteredExpandedState.Add(item, unfilteredState);
            }
        }

        private List<CodeExplorerItemViewModel> FlattenNodeList(IEnumerable<CodeExplorerItemViewModel> nodes)
        {
            var output = new List<CodeExplorerItemViewModel>();

            foreach (var item in nodes)
            {
                output.Add(item);
                if (item.Items.Any())
                {
                    output.AddRange(FlattenNodeList(item.Items));
                }                
            }

            return output;
        }

        private void ParserState_ModuleStateChanged(object sender, ParseProgressEventArgs e)
        {
            // if we are resolving references, we already have the declarations and don't need to display error
            if (!(e.State == ParserState.Error ||
                (e.State == ParserState.ResolverError &&
                e.OldState == ParserState.ResolvingDeclarations)))
            {
                return;
            }

            var componentProject = _state.ProjectsProvider.Project(e.Module.ProjectId);
            
            var projectNode = Projects.OfType<CodeExplorerProjectViewModel>()
                .FirstOrDefault(p => p.Declaration.Project?.Equals(componentProject) ?? false);

            if (projectNode == null)
            {
                return;
            }

            SetErrorState(projectNode, e.Module);

            if (_errorStateSet) { return; }

            // at this point, we know the node is newly added--we have to add a new node, not just change the icon of the old one.
            var projectName = componentProject.Name;
            var folderNode = projectNode.Items.FirstOrDefault(f => f is CodeExplorerCustomFolderViewModel && f.Name == projectName);

            _uiDispatcher.Invoke(() =>
            {
                try
                {
                    if (folderNode == null)
                    {
                        folderNode = new CodeExplorerCustomFolderViewModel(projectNode, projectName, projectName, _state.ProjectsProvider, _vbe);
                        projectNode.AddChild(folderNode);
                    }

                    var declaration = CreateDeclaration(e.Module);
                    var newNode =
                        new CodeExplorerComponentViewModel(folderNode, declaration, new List<Declaration>(), _state.ProjectsProvider, _vbe)
                        {
                            IsErrorState = true
                        };

                    folderNode.AddChild(newNode);

                    // Force a refresh. OnPropertyChanged("Projects") didn't work.
                    ForceProjectsRefresh(Projects);
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown trying to refresh the code explorer view on the UI thread.");
                }
            });
        }

        private Declaration CreateDeclaration(QualifiedModuleName module)
        {
            var projectDeclaration =
                _state.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                    .FirstOrDefault(item => item.ProjectId == module.ProjectId);

            if (module.ComponentType == ComponentType.StandardModule)
            {
                return new ProceduralModuleDeclaration(
                        new QualifiedMemberName(module, module.ComponentName), projectDeclaration,
                        module.ComponentName, true, new List<IAnnotation>(), null);
            }

            return new ClassModuleDeclaration(new QualifiedMemberName(module, module.ComponentName),
                    projectDeclaration, module.ComponentName, true, new List<IAnnotation>(), null);
        }

        private void ReorderChildNodes(IEnumerable<CodeExplorerItemViewModel> nodes)
        {
            foreach (var node in nodes)
            {
                node.ReorderItems(SortByName, GroupByType);
                ReorderChildNodes(node.Items);
            }
        }

        private bool _errorStateSet;
        private void SetErrorState(CodeExplorerItemViewModel itemNode, QualifiedModuleName module)
        {
            _errorStateSet = false;

            foreach (var node in itemNode.Items)
            {
                if (node is CodeExplorerCustomFolderViewModel)
                {
                    SetErrorState(node, module);
                }

                if (_errorStateSet)
                {
                    return;
                }

                var componentNode = node as CodeExplorerComponentViewModel;
                if (componentNode?.Declaration.QualifiedName.QualifiedModuleName.Equals(module) == true)
                {
                    componentNode.IsErrorState = true;
                    _errorStateSet = true;
                }
            }
        }

        private void ExecuteCollapseNodes(object parameter)
        {
            if (!(parameter is CodeExplorerItemViewModel node))
            {
                return;
            }

            SwitchNodeState(node, false);
        }

        private void ExecuteExpandNodes(object parameter)
        {
            if (!(parameter is CodeExplorerItemViewModel node))
            {
                return;
            }

            SwitchNodeState(node, true);
        }

        private void SwitchNodeState(CodeExplorerItemViewModel node, bool expandedState)
        {
            node.IsExpanded = expandedState;

            foreach (var item in node.Items)
            {
                item.IsExpanded = expandedState;
                SwitchNodeState(item, expandedState);
            }
        }

        /// <summary>
        /// Works backward from the passed node and expands all parents to make it visible.
        /// </summary>
        /// <param name="node"></param>
        private void ExpandToNode(CodeExplorerItemViewModel node)
        {
            while (true)
            {
                if (node == null)
                {
                    return;
                }
                node.IsExpanded = true;
                node = node.Parent;
            }
        }

        private string _filterText = string.Empty;
        private Dictionary<CodeExplorerItemViewModel, bool> _unfilteredExpandedState = new Dictionary<CodeExplorerItemViewModel, bool>();

        public string FilterText
        {
            get => _filterText;
            set
            {
                var input = value ?? string.Empty;
                if (_filterText.Equals(input))
                {
                    return;
                }

                if (string.IsNullOrEmpty(_filterText) && !string.IsNullOrEmpty(input))
                {
                    CacheUnfilteredState();
                }
                else if (string.IsNullOrEmpty(input) && !string.IsNullOrEmpty(_filterText))
                {
                    RestoreUnfilteredState();
                }

                _filterText = value;
                if (!string.IsNullOrEmpty(_filterText))
                {
                    FilterByName(Projects, _filterText);
                }
                
                OnPropertyChanged();
            }
        }

        private void CacheUnfilteredState()
        {
            _unfilteredExpandedState = FlattenNodeList(Projects).Where(node => node != null)
                .ToDictionary(node => node, node => node.IsExpanded);
        }

        private void RestoreUnfilteredState()
        {
            foreach (var node in FlattenNodeList(Projects).Where(node => node != null && _unfilteredExpandedState.ContainsKey(node)))
            {
                node.IsExpanded = _unfilteredExpandedState[node];
                node.IsVisible = true;
            }
        }

        private double? _fontSize = 10;
        public double? FontSize
        {
            get => _fontSize ?? 10;
            set
            {
                if (!value.HasValue || value.Equals(_fontSize))
                {
                    return;
                }

                _fontSize = value;
                OnPropertyChanged();
            }
        }
        
        public ReparseCommand RefreshCommand { get; set; }

        public OpenCommand OpenCommand { get; set; }

        public AddVBFormCommand AddVBFormCommand { get; set; }
        public AddMDIFormCommand AddMDIFormCommand { get; set; }
        public AddUserFormCommand AddUserFormCommand { get; set; }
        public AddStdModuleCommand AddStdModuleCommand { get; set; }
        public AddClassModuleCommand AddClassModuleCommand { get; set; }                
        public AddUserControlCommand AddUserControlCommand { get; set; }
        public AddPropertyPageCommand AddPropertyPageCommand { get; set; }
        public AddUserDocumentCommand AddUserDocumentCommand { get; set; }
        public AddTestComponentCommand AddTestModuleCommand { get; set; }
        public AddTestModuleWithStubsCommand AddTestModuleWithStubsCommand { get; set; }
		public AddTemplateCommand AddTemplateCommand { get; set; }
        public OpenDesignerCommand OpenDesignerCommand { get; set; }
        public OpenProjectPropertiesCommand OpenProjectPropertiesCommand { get; set; }
        public SetAsStartupProjectCommand SetAsStartupProjectCommand { get; set; }
        public RenameCommand RenameCommand { get; set; }
        public IndentCommand IndenterCommand { get; set; }
        public CodeExplorerFindAllReferencesCommand FindAllReferencesCommand { get; set; }
        public FindAllImplementationsCommand FindAllImplementationsCommand { get; set; }
        public CommandBase CollapseAllSubnodesCommand { get; }
        public CopyResultsCommand CopyResultsCommand { get; set; }
        public CommandBase ExpandAllSubnodesCommand { get; }
        public ImportCommand ImportCommand { get; set; }
        public ExportCommand ExportCommand { get; set; }
        public ExportAllCommand ExportAllCommand { get; set; }
        public CommandBase RemoveCommand { get; }
        public PrintCommand PrintCommand { get; set; }
        public AddRemoveReferencesCommand AddRemoveReferencesCommand { get; set; }
        public ICommand ClearFilterTextCommand { get; }
        public CommandBase SyncCodePaneCommand { get; }

        private readonly RemoveCommand _externalRemoveCommand;

        private static readonly Type[] ProjectNodes =
        {
            typeof(CodeExplorerProjectViewModel),
            typeof(CodeExplorerComponentViewModel),
            typeof(CodeExplorerMemberViewModel)
        };

        private static readonly List<DeclarationType> DeclarationsWithNodes =
            CodeExplorerProjectViewModel.ComponentTypes
                .Concat(CodeExplorerComponentViewModel.MemberTypes)
                .Concat(CodeExplorerMemberViewModel.SubMemberTypes).ToList();

        public CodeExplorerItemViewModel FindVisibleNodeForDeclaration(Declaration declaration)
        {
            if (declaration == null)
            {
                return null;
            }

            var visible = FlattenNodeList(Projects)
                .Where(node => node?.Declaration != null && 
                               ProjectNodes.Contains(node.GetType()) && 
                               node.IsVisible &&
                               node.Declaration.ProjectId.Equals(declaration.ProjectId));

            if (declaration.DeclarationType == DeclarationType.Project)
            {
                return visible.OfType<CodeExplorerProjectViewModel>().FirstOrDefault();
            }

            return DeclarationsWithNodes.Contains(declaration.DeclarationType)
                ? visible.FirstOrDefault(node => node.Declaration.DeclarationType == declaration.DeclarationType && 
                                                 node.Declaration.QualifiedName.Equals(declaration.QualifiedName))
                : null;
        }

        // this is a special case--we have to reset SelectedItem to prevent a crash
        private void ExecuteRemoveCommand(object param)
        {
            var node = (CodeExplorerComponentViewModel)SelectedItem;
            SelectedItem = Projects.FirstOrDefault(p => p.QualifiedSelection.HasValue 
                && p.QualifiedSelection.Value.QualifiedName.ProjectId == node.Declaration.ProjectId);

            _externalRemoveCommand.Execute(param);
        }

        private bool CanExecuteExportAllCommand => ExportAllCommand?.CanExecute(SelectedItem) ?? false;

        public Visibility ExportVisibility => _vbe.Kind == VBEKind.Standalone || CanExecuteExportAllCommand ? Visibility.Collapsed : Visibility.Visible;

        public Visibility ExportAllVisibility => CanExecuteExportAllCommand ? Visibility.Visible : Visibility.Collapsed;

        public Visibility TreeViewVisibility => Projects == null || Projects.Count == 0 ? Visibility.Collapsed : Visibility.Visible;

        public Visibility EmptyUIRefreshMessageVisibility => _isBusy ? Visibility.Hidden : Visibility.Visible;

        public Visibility VB6Visibility => _vbe.Kind == VBEKind.Standalone ? Visibility.Visible : Visibility.Collapsed;

        public Visibility VBAVisibility => _vbe.Kind == VBEKind.Hosted ? Visibility.Visible : Visibility.Collapsed;

        public void FilterByName(IEnumerable<CodeExplorerItemViewModel> nodes, string searchString)
        {
            if (string.IsNullOrEmpty(searchString))
            {
                return;      
            }

            foreach (var item in nodes)
            {
                if (item == null) { continue; }
                
                if (item.Items.Any())
                {
                    FilterByName(item.Items, searchString);
                }

                item.IsVisible = string.IsNullOrEmpty(searchString) ||
                                 item.Items.Any(c => c.IsVisible) ||
                                 item.Name.ToLowerInvariant().Contains(searchString.ToLowerInvariant());

                if (!string.IsNullOrEmpty(FilterText) && item.IsVisible && !item.IsExpanded)
                {
                    ExpandToNode(item);
                }
            }
        }

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= HandleStateChanged;
                _state.ModuleStateChanged -= ParserState_ModuleStateChanged;
            }
        }
    }
}

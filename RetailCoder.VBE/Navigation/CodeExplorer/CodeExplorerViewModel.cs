using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Windows;

// ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
// ReSharper disable ExplicitCallerInfoArgument

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerViewModel : ViewModelBase, IDisposable
    {
        private readonly FolderHelper _folderHelper;
        private readonly RubberduckParserState _state;
        private IConfigProvider<GeneralSettings> _generalSettingsProvider;
        private readonly IConfigProvider<WindowSettings> _windowSettingsProvider;
        private readonly GeneralSettings _generalSettings;
        private readonly WindowSettings _windowSettings;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public CodeExplorerViewModel(FolderHelper folderHelper, RubberduckParserState state, List<CommandBase> commands,
            IConfigProvider<GeneralSettings> generalSettingsProvider, IConfigProvider<WindowSettings> windowSettingsProvider)
        {
            _folderHelper = folderHelper;
            _state = state;
            _state.StateChanged += HandleStateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;
            _generalSettingsProvider = generalSettingsProvider;
            _windowSettingsProvider = windowSettingsProvider;

            if (generalSettingsProvider != null)
            {
                _generalSettings = generalSettingsProvider.Create();
            }

            if (windowSettingsProvider != null)
            {
                _windowSettings = windowSettingsProvider.Create();
            }

            var reparseCommand = commands.OfType<ReparseCommand>().SingleOrDefault();

            RefreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), 
                reparseCommand == null ? (Action<object>)(o => { }) :
                o => reparseCommand.Execute(o),
                o => !IsBusy && reparseCommand != null && reparseCommand.CanExecute(o));

            OpenCommand = commands.OfType<UI.CodeExplorer.Commands.OpenCommand>().SingleOrDefault();
            OpenDesignerCommand = commands.OfType<OpenDesignerCommand>().SingleOrDefault();

            AddTestModuleCommand = commands.OfType<UI.CodeExplorer.Commands.AddTestModuleCommand>().SingleOrDefault();
            AddTestModuleWithStubsCommand = commands.OfType<AddTestModuleWithStubsCommand>().SingleOrDefault();

            AddStdModuleCommand = commands.OfType<AddStdModuleCommand>().SingleOrDefault();
            AddClassModuleCommand = commands.OfType<AddClassModuleCommand>().SingleOrDefault();
            AddUserFormCommand = commands.OfType<AddUserFormCommand>().SingleOrDefault();

            OpenProjectPropertiesCommand = commands.OfType<OpenProjectPropertiesCommand>().SingleOrDefault();
            RenameCommand = commands.OfType<RenameCommand>().SingleOrDefault();
            IndenterCommand = commands.OfType<IndentCommand>().SingleOrDefault();

            FindAllReferencesCommand = commands.OfType<UI.CodeExplorer.Commands.FindAllReferencesCommand>().SingleOrDefault();
            FindAllImplementationsCommand = commands.OfType<UI.CodeExplorer.Commands.FindAllImplementationsCommand>().SingleOrDefault();

            CollapseAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCollapseNodes);
            ExpandAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteExpandNodes);

            ImportCommand = commands.OfType<ImportCommand>().SingleOrDefault();
            ExportCommand = commands.OfType<ExportCommand>().SingleOrDefault();
            ExportAllCommand = commands.OfType<Rubberduck.UI.Command.ExportAllCommand>().SingleOrDefault();
            
            _externalRemoveCommand = commands.OfType<RemoveCommand>().SingleOrDefault();
            if (_externalRemoveCommand != null)
            {
                RemoveCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveComand, _externalRemoveCommand.CanExecute);
            }

            PrintCommand = commands.OfType<PrintCommand>().SingleOrDefault();

            CommitCommand = commands.OfType<CommitCommand>().SingleOrDefault();
            UndoCommand = commands.OfType<UndoCommand>().SingleOrDefault();

            CopyResultsCommand = commands.OfType<CopyResultsCommand>().SingleOrDefault();

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
        }

        private CodeExplorerItemViewModel _selectedItem;
        public CodeExplorerItemViewModel SelectedItem
        {
            get => _selectedItem;
            set
            {
                _selectedItem = value;
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

        public CommandBase CopyResultsCommand { get; }

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

                if (!(SelectedItem is ICodeExplorerDeclarationViewModel))
                {
                    return SelectedItem.Name;
                }

                var declaration = SelectedItem.GetSelectedDeclaration();
                
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
                if (SelectedItem is ICodeExplorerDeclarationViewModel)
                {
                    return ((ICodeExplorerDeclarationViewModel)SelectedItem).Declaration.DescriptionString;
                }

                if (SelectedItem is CodeExplorerCustomFolderViewModel)
                {
                    return ((CodeExplorerCustomFolderViewModel)SelectedItem).FolderAttribute;
                }

                return string.Empty;
            }
        }

        public bool CanExecuteIndenterCommand => IndenterCommand.CanExecute(SelectedItem);
        public bool CanExecuteRenameCommand => RenameCommand.CanExecute(SelectedItem);
        public bool CanExecuteFindAllReferencesCommand => FindAllReferencesCommand.CanExecute(SelectedItem);

        private ObservableCollection<CodeExplorerItemViewModel> _projects;
        public ObservableCollection<CodeExplorerItemViewModel> Projects
        {
            get => _projects;
            set
            {
                ReorderChildNodes(value);
                _projects = new ObservableCollection<CodeExplorerItemViewModel>(value.OrderBy(o => o.NameWithSignature));
                CanSearch = _projects.Any();

                OnPropertyChanged();
                // Once a Project has been set, show the TreeView
                OnPropertyChanged("TreeViewVisibility");
                OnPropertyChanged("CanSearch");
            }
        }

        private void HandleStateChanged(object sender, ParserStateEventArgs e)
        {
            if (Projects == null)
            {
                Projects = new ObservableCollection<CodeExplorerItemViewModel>();
            }

            IsBusy = _state.Status != ParserState.Pending && _state.Status <= ParserState.ResolvedDeclarations;

            if (e.State != ParserState.ResolvedDeclarations)
            {
                return;
            }

            var userDeclarations = _state.DeclarationFinder.AllUserDeclarations
                .GroupBy(declaration => declaration.ProjectId)
                .ToList();

            if (userDeclarations.Any(
                    grouping => grouping.All(declaration => declaration.DeclarationType != DeclarationType.Project)))
            {
                return;
            }

            var newProjects = userDeclarations.Select(grouping =>
                new CodeExplorerProjectViewModel(_folderHelper,
                    grouping.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Project),
                    grouping)).ToList();

            UpdateNodes(Projects, newProjects);
            
            Projects = new ObservableCollection<CodeExplorerItemViewModel>(newProjects);
        }

        private void UpdateNodes(IEnumerable<CodeExplorerItemViewModel> oldList, IEnumerable<CodeExplorerItemViewModel> newList)
        {
            foreach (var item in newList)
            {
                CodeExplorerItemViewModel oldItem;

                if (item is CodeExplorerCustomFolderViewModel)
                {
                    oldItem = oldList.FirstOrDefault(i => i.Name == item.Name);
                }
                else
                {
                    oldItem = oldList.FirstOrDefault(i =>
                        item.QualifiedSelection != null && i.QualifiedSelection != null &&
                        i.QualifiedSelection.Value.QualifiedName.ProjectId ==
                        item.QualifiedSelection.Value.QualifiedName.ProjectId &&
                        i.QualifiedSelection.Value.QualifiedName.ComponentName ==
                        item.QualifiedSelection.Value.QualifiedName.ComponentName &&
                        i.QualifiedSelection.Value.Selection == item.QualifiedSelection.Value.Selection);
                }

                if (oldItem != null)
                {
                    item.IsExpanded = oldItem.IsExpanded;
                    item.IsSelected = oldItem.IsSelected;

                    if (oldItem.Items.Any() && item.Items.Any())
                    {
                        UpdateNodes(oldItem.Items, item.Items);
                    }
                }
            }
        }

        private void ParserState_ModuleStateChanged(object sender, Parsing.ParseProgressEventArgs e)
        {
            // if we are resolving references, we already have the declarations and don't need to display error
            if (!(e.State == ParserState.Error ||
                (e.State == ParserState.ResolverError &&
                e.OldState == ParserState.ResolvingDeclarations)))
            {
                return;
            }

            var components = e.Module.Component.Collection;
            var componentProject = components.Parent;
            {
                var projectNode = Projects.OfType<CodeExplorerProjectViewModel>()
                    .FirstOrDefault(p => p.Declaration.Project.Equals(componentProject));

                if (projectNode == null)
                {
                    return;
                }

                SetErrorState(projectNode, e.Module);

                if (_errorStateSet) { return; }

                // at this point, we know the node is newly added--we have to add a new node, not just change the icon of the old one.
                var projectName = componentProject.Name;
                var folderNode = projectNode.Items.FirstOrDefault(f => f is CodeExplorerCustomFolderViewModel && f.Name == projectName);

                UiDispatcher.Invoke(() =>
                {
                    try
                    {
                        if (folderNode == null)
                        {
                            folderNode = new CodeExplorerCustomFolderViewModel(projectNode, projectName, projectName);
                            projectNode.AddChild(folderNode);
                        }

                        var declaration = CreateDeclaration(e.Module);
                        var newNode =
                            new CodeExplorerComponentViewModel(folderNode, declaration, new List<Declaration>())
                            {
                                IsErrorState = true
                            };

                        folderNode.AddChild(newNode);

                        // Force a refresh. OnPropertyChanged("Projects") didn't work.
                        Projects = Projects;
                    }
                    catch (Exception exception)
                    {
                        Logger.Error(exception, "Exception thrown trying to refresh the code explorer view on the UI thread.");
                    }
                });
            }
        }

        private Declaration CreateDeclaration(QualifiedModuleName module)
        {
            var projectDeclaration =
                _state.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                    .FirstOrDefault(item => item.Project.VBComponents.Contains(module.Component));

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
                if (componentNode?.GetSelectedDeclaration().QualifiedName.QualifiedModuleName.Equals(module) == true)
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

        public CommandBase RefreshCommand { get; }

        public CommandBase OpenCommand { get; }

        public CommandBase AddTestModuleCommand { get; }
        public CommandBase AddTestModuleWithStubsCommand { get; }
        public CommandBase AddStdModuleCommand { get; }
        public CommandBase AddClassModuleCommand { get; }
        public CommandBase AddUserFormCommand { get; }

        public CommandBase OpenDesignerCommand { get; }
        public CommandBase OpenProjectPropertiesCommand { get; }

        public CommandBase RenameCommand { get; }

        public CommandBase IndenterCommand { get; }

        public CommandBase FindAllReferencesCommand { get; }
        public CommandBase FindAllImplementationsCommand { get; }

        public CommandBase CollapseAllSubnodesCommand { get; }
        public CommandBase ExpandAllSubnodesCommand { get; }

        public CommandBase ImportCommand { get; }
        public CommandBase ExportCommand { get; }
        public CommandBase ExportAllCommand { get; }

        public CommandBase RemoveCommand { get; }

        public CommandBase PrintCommand { get; }

        public CommandBase CommitCommand { get; }
        public CommandBase UndoCommand { get; }
        
        private readonly CommandBase _externalRemoveCommand;

        // this is a special case--we have to reset SelectedItem to prevent a crash
        private void ExecuteRemoveComand(object param)
        {
            var node = (CodeExplorerComponentViewModel)SelectedItem;
            SelectedItem = Projects.FirstOrDefault(p => p.QualifiedSelection.HasValue 
                && p.QualifiedSelection.Value.QualifiedName.ProjectId == node.Declaration.ProjectId);

            _externalRemoveCommand.Execute(param);
        }

        private bool CanExecuteExportAllCommand => ExportAllCommand.CanExecute(SelectedItem);

        public Visibility ExportVisibility => CanExecuteExportAllCommand ? Visibility.Collapsed : Visibility.Visible;

        public Visibility ExportAllVisibility => CanExecuteExportAllCommand ? Visibility.Visible : Visibility.Collapsed;

        public bool EnableSourceControl => _generalSettings.EnableExperimentalFeatures.Any(a => a.Key == RubberduckUI.GeneralSettings_EnableSourceControl);

        public Visibility TreeViewVisibility => Projects == null || Projects.Count == 0 ? Visibility.Collapsed : Visibility.Visible;

        public Visibility EmptyUIRefreshMessageVisibility => _isBusy ? Visibility.Hidden : Visibility.Visible;

        public void FilterByName(IEnumerable<CodeExplorerItemViewModel> nodes, string searchString)
        {
            foreach (var item in nodes)
            {
                if (item == null) { continue; }
                
                if (item.Items.Any())
                {
                    FilterByName(item.Items, searchString);
                }
                
                item.IsVisible = item.Items.Any(c => c.IsVisible) ||
                                 item.Name.ToLowerInvariant().Contains(searchString.ToLowerInvariant()) ||
                                 string.IsNullOrEmpty(searchString);
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

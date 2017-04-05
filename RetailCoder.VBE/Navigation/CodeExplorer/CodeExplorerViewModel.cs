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
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable CanBeReplacedWithTryCastAndCheckForNull

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerViewModel : ViewModelBase, IDisposable
    {
        private readonly FolderHelper _folderHelper;
        private readonly RubberduckParserState _state;

        public CodeExplorerViewModel(FolderHelper folderHelper, RubberduckParserState state, List<CommandBase> commands)
        {
            _folderHelper = folderHelper;
            _state = state;
            _state.StateChanged += HandleStateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;

            var reparseCommand = commands.OfType<ReparseCommand>().SingleOrDefault();

            RefreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), 
                reparseCommand == null ? (Action<object>)(o => { }) :
                o => reparseCommand.Execute(o),
                o => !IsBusy && reparseCommand != null && reparseCommand.CanExecute(o));
            
            NavigateCommand = commands.OfType<UI.CodeExplorer.Commands.NavigateCommand>().SingleOrDefault();

            AddTestModuleCommand = commands.OfType<UI.CodeExplorer.Commands.AddTestModuleCommand>().SingleOrDefault();
            AddStdModuleCommand = commands.OfType<AddStdModuleCommand>().SingleOrDefault();
            AddClassModuleCommand = commands.OfType<AddClassModuleCommand>().SingleOrDefault();
            AddUserFormCommand = commands.OfType<AddUserFormCommand>().SingleOrDefault();

            OpenDesignerCommand = commands.OfType<OpenDesignerCommand>().SingleOrDefault();
            OpenProjectPropertiesCommand = commands.OfType<OpenProjectPropertiesCommand>().SingleOrDefault();
            RenameCommand = commands.OfType<RenameCommand>().SingleOrDefault();
            IndenterCommand = commands.OfType<IndentCommand>().SingleOrDefault();

            FindAllReferencesCommand = commands.OfType<UI.CodeExplorer.Commands.FindAllReferencesCommand>().SingleOrDefault();
            FindAllImplementationsCommand = commands.OfType<UI.CodeExplorer.Commands.FindAllImplementationsCommand>().SingleOrDefault();

            CollapseAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCollapseNodes);
            ExpandAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteExpandNodes);

            ImportCommand = commands.OfType<ImportCommand>().SingleOrDefault();
            ExportCommand = commands.OfType<ExportCommand>().SingleOrDefault();
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
                SortByName = (bool)param;
                SortBySelection = !(bool)param;
            });

            SetSelectionSortCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                SortBySelection = (bool)param;
                SortByName = !(bool)param;
            });
        }

        private CodeExplorerItemViewModel _selectedItem;
        public CodeExplorerItemViewModel SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                OnPropertyChanged();

                // ReSharper disable ExplicitCallerInfoArgument
                OnPropertyChanged("CanExecuteIndenterCommand");
                OnPropertyChanged("CanExecuteRenameCommand");
                OnPropertyChanged("CanExecuteFindAllReferencesCommand");
                OnPropertyChanged("PanelTitle");
                OnPropertyChanged("Description");
                // ReSharper restore ExplicitCallerInfoArgument
            }
        }

        private bool _sortByName = true;
        public bool SortByName
        {
            get { return _sortByName; }
            set
            {
                if (_sortByName == value)
                {
                    return;
                }

                _sortByName = value;
                OnPropertyChanged();

                ReorderChildNodes(Projects);
            }
        }

        private bool _sortBySelection;
        public bool SortBySelection
        {
            get { return _sortBySelection; }
            set
            {
                if (_sortBySelection == value)
                {
                    return;
                }

                _sortBySelection = value;
                OnPropertyChanged();

                ReorderChildNodes(Projects);
            }
        }

        public CommandBase CopyResultsCommand { get; }

        public CommandBase SetNameSortCommand { get; }

        public CommandBase SetSelectionSortCommand { get; }

        private bool _sortByType = true;
        public bool SortByType
        {
            get { return _sortByType; }
            set
            {
                if (_sortByType != value)
                {
                    _sortByType = value;
                    OnPropertyChanged();

                    ReorderChildNodes(Projects);
                }
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value;
                OnPropertyChanged();
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
            get { return _projects; }
            set
            {
                ReorderChildNodes(value);
                _projects = new ObservableCollection<CodeExplorerItemViewModel>(value.OrderBy(o => o.NameWithSignature));
                
                OnPropertyChanged();
            }
        }

        private void HandleStateChanged(object sender, ParserStateEventArgs e)
        {
            if (Projects == null)
            {
                Projects = new ObservableCollection<CodeExplorerItemViewModel>();
            }

            IsBusy = _state.Status != ParserState.Pending && _state.Status < ParserState.ResolvedDeclarations;
            if (e.State != ParserState.ResolvedDeclarations)
            {
                return;
            }

            var userDeclarations = _state.AllUserDeclarations
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

            var components = e.Component.Collection;
            var componentProject = components.Parent;
            {
                var projectNode = Projects.OfType<CodeExplorerProjectViewModel>()
                    .FirstOrDefault(p => p.Declaration.Project.Equals(componentProject));

                if (projectNode == null)
                {
                    return;
                }

                SetErrorState(projectNode, e.Component);

                if (_errorStateSet) { return; }

                // at this point, we know the node is newly added--we have to add a new node, not just change the icon of the old one.
                var projectName = componentProject.Name;
                var folderNode = projectNode.Items.FirstOrDefault(f => f is CodeExplorerCustomFolderViewModel && f.Name == projectName);

                UiDispatcher.Invoke(() =>
                {
                    if (folderNode == null)
                    {
                        folderNode = new CodeExplorerCustomFolderViewModel(projectNode, projectName, projectName);
                        projectNode.AddChild(folderNode);
                    }

                    var declaration = CreateDeclaration(e.Component);
                    var newNode = new CodeExplorerComponentViewModel(folderNode, declaration, new List<Declaration>())
                    {
                        IsErrorState = true
                    };

                    folderNode.AddChild(newNode);

                    // Force a refresh. OnPropertyChanged("Projects") didn't work.
                    Projects = Projects;
                });
            }
        }

        private Declaration CreateDeclaration(IVBComponent component)
        {
            var projectDeclaration =
                _state.AllUserDeclarations.FirstOrDefault(item =>
                        item.DeclarationType == DeclarationType.Project &&
                        item.Project.VBComponents.Contains(component));

            if (component.Type == ComponentType.StandardModule)
            {
                return new ProceduralModuleDeclaration(
                        new QualifiedMemberName(new QualifiedModuleName(component), component.Name), projectDeclaration,
                        component.Name, true, new List<IAnnotation>(), null);
            }

            return new ClassModuleDeclaration(new QualifiedMemberName(new QualifiedModuleName(component), component.Name),
                    projectDeclaration, component.Name, true, new List<IAnnotation>(), null);
        }

        private void ReorderChildNodes(IEnumerable<CodeExplorerItemViewModel> nodes)
        {
            foreach (var node in nodes)
            {
                node.ReorderItems(SortByName, SortByType);
                ReorderChildNodes(node.Items);
            }
        }

        private bool _errorStateSet;
        private void SetErrorState(CodeExplorerItemViewModel itemNode, IVBComponent component)
        {
            _errorStateSet = false;

            foreach (var node in itemNode.Items)
            {
                if (node is CodeExplorerCustomFolderViewModel)
                {
                    SetErrorState(node, component);
                }

                if (_errorStateSet)
                {
                    return;
                }

                var componentNode = node as CodeExplorerComponentViewModel;
                if (componentNode?.GetSelectedDeclaration().QualifiedName.QualifiedModuleName.Component.Equals(component) == true)
                {
                    componentNode.IsErrorState = true;
                    _errorStateSet = true;
                }
            }
        }

        private void ExecuteCollapseNodes(object parameter)
        {
            var node = parameter as CodeExplorerItemViewModel;
            if (node == null) { return; }

            SwitchNodeState(node, false);
        }

        private void ExecuteExpandNodes(object parameter)
        {
            var node = parameter as CodeExplorerItemViewModel;
            if (node == null) { return; }

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

        public CommandBase NavigateCommand { get; }

        public CommandBase AddTestModuleCommand { get; }
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

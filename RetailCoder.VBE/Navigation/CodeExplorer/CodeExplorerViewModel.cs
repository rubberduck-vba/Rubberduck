using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;

// ReSharper disable CanBeReplacedWithTryCastAndCheckForNull

namespace Rubberduck.Navigation.CodeExplorer
{
    public sealed class CodeExplorerViewModel : ViewModelBase, IDisposable
    {
        private readonly FolderHelper _folderHelper;
        private readonly RubberduckParserState _state;
        private readonly Dispatcher _dispatcher;

        public CodeExplorerViewModel(FolderHelper folderHelper, RubberduckParserState state, List<ICommand> commands)
        {
            _dispatcher = Dispatcher.CurrentDispatcher;

            _folderHelper = folderHelper;
            _state = state;
            _state.StateChanged += ParserState_StateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;

            _refreshCommand = commands.OfType<CodeExplorer_RefreshCommand>().FirstOrDefault();
            _refreshComponentCommand = commands.OfType<CodeExplorer_RefreshComponentCommand>().FirstOrDefault();
            _navigateCommand = commands.OfType<CodeExplorer_NavigateCommand>().FirstOrDefault();

            _addTestModuleCommand = commands.OfType<CodeExplorer_AddTestModuleCommand>().FirstOrDefault();
            _addStdModuleCommand = commands.OfType<CodeExplorer_AddStdModuleCommand>().FirstOrDefault();
            _addClassModuleCommand = commands.OfType<CodeExplorer_AddClassModuleCommand>().FirstOrDefault();
            _addUserFormCommand = commands.OfType<CodeExplorer_AddUserFormCommand>().FirstOrDefault();

            _openDesignerCommand = commands.OfType<CodeExplorer_OpenDesignerCommand>().FirstOrDefault();
            _renameCommand = commands.OfType<CodeExplorer_RenameCommand>().FirstOrDefault();
            _indenterCommand = commands.OfType<CodeExplorer_IndentCommand>().FirstOrDefault();

            _findAllReferencesCommand = commands.OfType<CodeExplorer_FindAllReferencesCommand>().FirstOrDefault();
            _findAllImplementationsCommand = commands.OfType<CodeExplorer_FindAllImplementationsCommand>().FirstOrDefault();

            _importCommand = commands.OfType<CodeExplorer_ImportCommand>().FirstOrDefault();
            _exportCommand = commands.OfType<CodeExplorer_ExportCommand>().FirstOrDefault();
            _externalRemoveCommand = commands.OfType<CodeExplorer_RemoveCommand>().FirstOrDefault();
            if (_externalRemoveCommand != null)
            {
                _removeCommand = new DelegateCommand(ExecuteRemoveComand, _externalRemoveCommand.CanExecute);
            }

            _printCommand = commands.OfType<CodeExplorer_PrintCommand>().FirstOrDefault();

            _commitCommand = commands.OfType<CodeExplorer_CommitCommand>().FirstOrDefault();
            _undoCommand = commands.OfType<CodeExplorer_UndoCommand>().FirstOrDefault();
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
                CanRefresh = !_isBusy;
            }
        }

        private bool _canRefresh = true;
        public bool CanRefresh
        {
            get { return _canRefresh; }
            private set
            {
                _canRefresh = value;
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

                if (SelectedItem is CodeExplorerProjectViewModel)
                {
                    var node = (CodeExplorerProjectViewModel) SelectedItem;
                    return node.Declaration.IdentifierName + string.Format(" - ({0})", node.Declaration.DeclarationType);
                }

                if (SelectedItem is CodeExplorerComponentViewModel)
                {
                    var node = (CodeExplorerComponentViewModel) SelectedItem;
                    return node.Declaration.IdentifierName + string.Format(" - ({0})", node.Declaration.DeclarationType);
                }

                if (SelectedItem is CodeExplorerMemberViewModel)
                {
                    var node = (CodeExplorerMemberViewModel) SelectedItem;
                    return node.Declaration.IdentifierName + string.Format(" - ({0})", node.Declaration.DeclarationType);
                }

                return SelectedItem.Name;
            }
        }
        
        public string Description
        {
            get
            {
                if (SelectedItem is ICodeExplorerDeclarationViewModel)
                {
                    return ((ICodeExplorerDeclarationViewModel) SelectedItem).Declaration.DescriptionString;
                }

                if (SelectedItem is CodeExplorerCustomFolderViewModel)
                {
                    return ((CodeExplorerCustomFolderViewModel) SelectedItem).FolderAttribute;
                }

                if (SelectedItem is CodeExplorerErrorNodeViewModel)
                {
                    return ((CodeExplorerErrorNodeViewModel) SelectedItem).Name;
                }

                return string.Empty;
            }
        }

        public bool CanExecuteIndenterCommand { get { return IndenterCommand.CanExecute(SelectedItem); } }
        public bool CanExecuteRenameCommand { get { return RenameCommand.CanExecute(SelectedItem); } }
        public bool CanExecuteFindAllReferencesCommand { get { return FindAllReferencesCommand.CanExecute(SelectedItem); } }

        private ObservableCollection<CodeExplorerItemViewModel> _projects;
        public ObservableCollection<CodeExplorerItemViewModel> Projects
        {
            get { return _projects; }
            set
            {
                _projects = new ObservableCollection<CodeExplorerItemViewModel>(value.OrderBy(o => o.NameWithSignature));
                
                ReorderChildNodes(_projects);
                OnPropertyChanged();
            }
        }

        private void ParserState_StateChanged(object sender, EventArgs e)
        {
            if (Projects == null)
            {
                Projects = new ObservableCollection<CodeExplorerItemViewModel>();
            }

            IsBusy = _state.Status == ParserState.Parsing;
            if (_state.Status != ParserState.Ready)
            {
                return;
            }
            
            var userDeclarations = _state.AllUserDeclarations
                .GroupBy(declaration => declaration.Project)
                .Where(grouping => grouping.Key != null)
                .ToList();

            if (userDeclarations.Any(
                    grouping => grouping.All(declaration => declaration.DeclarationType != DeclarationType.Project)))
            {
                return;
            }


            var components = userDeclarations.SelectMany(s => s.Key.VBComponents.Cast<VBComponent>()).FirstOrDefault(s => s.Name == "asdf");

            var newProjects = userDeclarations.Select(grouping =>
                new CodeExplorerProjectViewModel(_folderHelper,
                    grouping.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Project),
                    grouping)).ToList();

            UpdateNodes(Projects, newProjects);
            
            Projects = new ObservableCollection<CodeExplorerItemViewModel>(newProjects);
        }

        private void UpdateNodes(IEnumerable<CodeExplorerItemViewModel> oldList,
            IEnumerable<CodeExplorerItemViewModel> newList)
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

                    if (oldItem.Items.Any() && item.Items.Any())
                    {
                        UpdateNodes(oldItem.Items, item.Items);
                    }
                }
            }
        }

        private void ParserState_ModuleStateChanged(object sender, Parsing.ParseProgressEventArgs e)
        {
            if (e.State != ParserState.Error)
            {
                return;
            }

            var componentProject = e.Component.Collection.Parent;
            var node = Projects.OfType<CodeExplorerProjectViewModel>()
                .FirstOrDefault(p => p.Declaration.Project == componentProject);

            if (node == null)
            {
                return;
            }

            var folderNode = node.Items.First(f => f is CodeExplorerCustomFolderViewModel && f.Name == node.Name);

            AddErrorNode addNode = AddComponentErrorNode;
            _dispatcher.BeginInvoke(addNode, node, folderNode, e.Component.Name);
        }

        private delegate void AddErrorNode(CodeExplorerItemViewModel projectNode, CodeExplorerItemViewModel folderNode, string componentName);
        private void AddComponentErrorNode(CodeExplorerItemViewModel projectNode, CodeExplorerItemViewModel folderNode,
            string componentName)
        {
            Projects.Remove(projectNode);
            RemoveFailingComponent(projectNode, componentName);

            folderNode.AddChild(new CodeExplorerErrorNodeViewModel(folderNode, componentName));
            Projects.Add(projectNode);

            if (SortByName)
            {
                // this setter does not ensure the values are the same
                // it also sorts the projects by name
                Projects = Projects;
            }
        }

        private void ReorderChildNodes(IEnumerable<CodeExplorerItemViewModel> nodes)
        {
            foreach (var node in nodes)
            {
                node.ReorderItems(SortByName, SortByType);
                ReorderChildNodes(node.Items);
            }
        }

        private bool _removedNode;
        private void RemoveFailingComponent(CodeExplorerItemViewModel itemNode, string componentName)
        {
            foreach (var node in itemNode.Items)
            {
                if (node is CodeExplorerCustomFolderViewModel)
                {
                    RemoveFailingComponent(node, componentName);
                }

                if (_removedNode)
                {
                    return;
                }

                if (node is CodeExplorerComponentViewModel)
                {
                    var component = (CodeExplorerComponentViewModel) node;
                    if (component.Name == componentName)
                    {
                        itemNode.Items.Remove(node);
                        _removedNode = true;

                        return;
                    }
                }
            }
        }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

        private readonly ICommand _refreshComponentCommand;
        public ICommand RefreshComponentCommand { get { return _refreshComponentCommand; } }

        private readonly ICommand _navigateCommand;
        public ICommand NavigateCommand { get { return _navigateCommand; } }

        private readonly ICommand _addTestModuleCommand;
        public ICommand AddTestModuleCommand { get { return _addTestModuleCommand; } }

        private readonly ICommand _addStdModuleCommand;
        public ICommand AddStdModuleCommand { get { return _addStdModuleCommand; } }

        private readonly ICommand _addClassModuleCommand;
        public ICommand AddClassModuleCommand { get { return _addClassModuleCommand; } }

        private readonly ICommand _addUserFormCommand;
        public ICommand AddUserFormCommand { get { return _addUserFormCommand; } }

        private readonly ICommand _openDesignerCommand;
        public ICommand OpenDesignerCommand { get { return _openDesignerCommand; } }

        private readonly ICommand _renameCommand;
        public ICommand RenameCommand { get { return _renameCommand; } }

        private readonly ICommand _indenterCommand;
        public ICommand IndenterCommand { get { return _indenterCommand; } }

        private readonly ICommand _findAllReferencesCommand;
        public ICommand FindAllReferencesCommand { get { return _findAllReferencesCommand; } }

        private readonly ICommand _findAllImplementationsCommand;
        public ICommand FindAllImplementationsCommand { get { return _findAllImplementationsCommand; } }

        private readonly ICommand _importCommand;
        public ICommand ImportCommand { get { return _importCommand; } }

        private readonly ICommand _exportCommand;
        public ICommand ExportCommand { get { return _exportCommand; } }

        private readonly ICommand _removeCommand;
        public ICommand RemoveCommand { get { return _removeCommand; } }

        private readonly ICommand _printCommand;
        public ICommand PrintCommand { get { return _printCommand; } }

        private readonly ICommand _commitCommand;
        public ICommand CommitCommand { get { return _commitCommand; } }

        private readonly ICommand _undoCommand;
        public ICommand UndoCommand { get { return _undoCommand; } }

        private readonly ICommand _externalRemoveCommand;

        // this is a special case--we have to reset SelectedItem to prevent a crash
        private void ExecuteRemoveComand(object param)
        {
            var node = (CodeExplorerComponentViewModel) SelectedItem;
            SelectedItem = Projects.First(p => ((CodeExplorerProjectViewModel) p).Declaration.Project == node.Declaration.Project);

            _externalRemoveCommand.Execute(param);
        }

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= ParserState_StateChanged;
                _state.ModuleStateChanged -= ParserState_ModuleStateChanged;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Windows;
using System.Windows.Input;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Templates;
using Rubberduck.UI.CodeExplorer.Commands.DragAndDrop;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.UnitTesting.ComCommands;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.CodeExplorer
{
    [Flags]
    public enum CodeExplorerSortOrder
    {
        Undefined = 0,
        Name = 1,
        CodeLine = 1 << 1,
        DeclarationType = 1 << 2,
        DeclarationTypeThenName = DeclarationType | Name,
        DeclarationTypeThenCodeLine = DeclarationType | CodeLine
    }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public sealed class CodeExplorerViewModel : ViewModelBase
    {
        // ReSharper disable NotAccessedField.Local - The settings providers aren't used, but several enhancement requests will need them.
        private readonly RubberduckParserState _state;
        private readonly RemoveCommand _externalRemoveCommand;
        private readonly IConfigurationService<GeneralSettings> _generalSettingsProvider;      
        private readonly IConfigurationService<WindowSettings> _windowSettingsProvider;
        private readonly IUiDispatcher _uiDispatcher;
        private readonly IVBE _vbe;
        private readonly ITemplateProvider _templateProvider;
        // ReSharper restore NotAccessedField.Local

        public CodeExplorerViewModel(
            RubberduckParserState state,
            RemoveCommand removeCommand,
            IConfigurationService<GeneralSettings> generalSettingsProvider, 
            IConfigurationService<WindowSettings> windowSettingsProvider, 
            IUiDispatcher uiDispatcher,
            IVBE vbe,
            ITemplateProvider templateProvider,
            ICodeExplorerSyncProvider syncProvider,
            IEnumerable<IAnnotation> annotations)
        {
            _state = state;
            _state.StateChanged += HandleStateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;

            _generalSettingsProvider = generalSettingsProvider;
            _generalSettingsProvider.SettingsChanged += GeneralSettingsChanged;
            RefreshDragAndDropSetting();

            _windowSettingsProvider = windowSettingsProvider;
            _uiDispatcher = uiDispatcher;
            _vbe = vbe;
            _templateProvider = templateProvider;
            _externalRemoveCommand = removeCommand;
            Annotations = annotations.ToList();

            CollapseAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCollapseNodes, EvaluateCanSwitchNodeState);
            ExpandAllSubnodesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteExpandNodes, EvaluateCanSwitchNodeState);
            ClearSearchCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteClearSearchCommand);
            if (_externalRemoveCommand != null)
            {
                RemoveCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveCommand, _externalRemoveCommand.CanExecute);
            }


            OnPropertyChanged(nameof(Projects));

            SyncCodePaneCommand = syncProvider.GetSyncCommand(this);
            // Force a call to EvaluateCanExecute
            OnPropertyChanged(nameof(SyncCodePaneCommand));
        }

        public ObservableCollection<ICodeExplorerNode> Projects { get; } = new ObservableCollection<ICodeExplorerNode>();

        public ObservableCollection<Template> BuiltInTemplates =>
            new ObservableCollection<Template>(_templateProvider.GetTemplates().Where(t => !t.IsUserDefined)
                .OrderBy(t => t.Name));

        public ObservableCollection<Template> UserDefinedTemplates =>
            new ObservableCollection<Template>(_templateProvider.GetTemplates().Where(t => t.IsUserDefined)
                .OrderBy(t => t.Name));

        public IEnumerable<IAnnotation> Annotations { get; }

        private ICodeExplorerNode _selectedItem;
        public ICodeExplorerNode SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (_selectedItem == value)
                {
                    return;
                }

                ExpandToNode(value);
                _selectedItem = value;

                OnPropertyChanged();

                OnPropertyChanged(nameof(ExportVisibility));
                OnPropertyChanged(nameof(ExportAllVisibility));
                OnPropertyChanged(nameof(CanBeAnnotated));
                OnPropertyChanged(nameof(AnyTemplatesCanExecute));
            }
        }

        public bool AnyTemplatesCanExecute =>
            AddTemplateCommand.CanExecuteForNode(SelectedItem)
            && BuiltInTemplates.Concat(UserDefinedTemplates)
                .Any(template => AddTemplateCommand.CanExecute((template.Name, SelectedItem)));

        public bool CanBeAnnotated =>
            AnnotateDeclarationCommand.CanExecuteForNode(SelectedItem)
            && Annotations.Any(annotation => AnnotateDeclarationCommand.CanExecute((annotation, SelectedItem)));

        private CodeExplorerSortOrder _sortOrder = CodeExplorerSortOrder.Name;
        public CodeExplorerSortOrder SortOrder
        {
            get => _sortOrder;
            set
            {
                if (_sortOrder == value)
                {
                    return;
                }

                _sortOrder = value;

                foreach (var project in Projects.OfType<CodeExplorerProjectViewModel>())
                {
                    project.SortOrder = _sortOrder;
                }

                OnPropertyChanged();
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
            }
        }

        private bool _unparsed = true;
        public bool Unparsed
        {
            get => _unparsed;
            set
            {
                if (_unparsed == value)
                {
                    return;
                }
                _unparsed = value;
                OnPropertyChanged();
            }
        }

        private string _filterText = string.Empty;
        public string Search
        {
            get => _filterText;
            set
            {
                var input = value ?? string.Empty;
                if (_filterText.Equals(input))
                {
                    return;
                }

                _filterText = value;

                foreach (var project in Projects)
                {
                    project.Filter = _filterText;
                }

                OnPropertyChanged();
            }
        }

        private double? _fontSize = 12;
        public double? FontSize
        {
            get => _fontSize ?? 12;
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

        // This is just a binding hack to force the icon binding to re-evaluate. It has no other functionality and should
        // not be used for anything else.
        public bool ParserReady
        {
            get => true;
            // ReSharper disable once ValueParameterNotUsed
            set => OnPropertyChanged();
        }

        private void HandleStateChanged(object sender, ParserStateEventArgs e)
        {
            Unparsed = false;

            if (e.State == ParserState.Ready && e.OldState != ParserState.Busy)
            {
                // Finished up resolving references, so we can now update the reference nodes.
                //We have to wait for the task to guarantee that no new parse starts invalidating all cached components.
                //CAUTION: This must not be executed from the UI thread!!!
                _uiDispatcher.StartTask(() =>
                {
                    var referenceFolders = Projects.SelectMany(node =>
                        node.Children.OfType<CodeExplorerReferenceFolderViewModel>());
                    foreach (var library in referenceFolders)
                    {
                        library.UpdateChildren();
                    }

                    Unparsed = !Projects.Any();
                    IsBusy = false;
                    ParserReady = true;
                }).Wait();
                return;
            }

            IsBusy = _state.Status != ParserState.Pending && _state.Status <= ParserState.ResolvedDeclarations;

            if (e.State == ParserState.ResolvedDeclarations)
            {
                Synchronize(_state.DeclarationFinder.AllUserDeclarations);
            }
        }

        /// <summary>
        /// Updates the ViewModel tree to reflect changes in user declarations after a reparse.
        /// </summary>
        /// <param name="declarations">
        /// The new declarations. This should always be the complete declaration set, and materializing
        /// the IEnumerable should be deferred to UI thread.
        /// </param>
        private void Synchronize(IEnumerable<Declaration> declarations)
        {
            //We have to wait for the task to guarantee that no new parse starts invalidating all cached components.
            _uiDispatcher.StartTask(() =>
            {
                var updates = declarations.ToList();
                var existing = Projects.OfType<CodeExplorerProjectViewModel>().ToList();

                foreach (var project in existing)
                {
                    project.Synchronize(ref updates);
                    if (project.Declaration is null)
                    {
                        Projects.Remove(project);
                    }
                }

                var adding = updates.OfType<ProjectDeclaration>().ToList();

                foreach (var project in adding)
                {
                    var model = new CodeExplorerProjectViewModel(project, ref updates, _state, _vbe, _state.ProjectsProvider) { Filter = Search };
                    Projects.Add(model);
                }

                CanSearch = Projects.Any();
            }).Wait();
        }

        private void ParserState_ModuleStateChanged(object sender, ParseProgressEventArgs e)
        {
            // if we are resolving references, we already have the declarations and don't need to display error
            if (!(e.State == ParserState.Error ||
                e.State == ParserState.ResolverError &&
                e.OldState == ParserState.ResolvingDeclarations))
            {
                return;
            }

            var componentProjectId = e.Module.ProjectId;
            
            var module = Projects.OfType<CodeExplorerProjectViewModel>()
                .FirstOrDefault(p => p.Declaration?.ProjectId.Equals(componentProjectId) ?? false)?.Children
                .OfType<CodeExplorerComponentViewModel>()
                .FirstOrDefault(component => component.QualifiedSelection?.QualifiedName.Equals(e.Module) ?? false);

            if (module == null)
            {
                return;
            }

            module.IsErrorState = true;
        }

        private void ExecuteClearSearchCommand(object parameter)
        {
            if (!string.IsNullOrEmpty(Search))
            {
                Search = string.Empty;
            }
        }

        private bool EvaluateCanSwitchNodeState(object parameter)
        {
            return SelectedItem?.Children?.Any() ?? false;
        }

        private void ExecuteCollapseNodes(object parameter)
        {
            if (!(parameter is ICodeExplorerNode node))
            {
                return;
            }

            SwitchNodeState(node, false);
        }

        private void ExecuteExpandNodes(object parameter)
        {
            if (!(parameter is ICodeExplorerNode node))
            {
                return;
            }

            SwitchNodeState(node, true);
        }

        // this is a special case--we have to reset SelectedItem to prevent a crash
        private void ExecuteRemoveCommand(object param)
        {
            var node = (CodeExplorerComponentViewModel)SelectedItem;
            SelectedItem = Projects.FirstOrDefault(p => p.QualifiedSelection.HasValue
                                                        && p.QualifiedSelection.Value.QualifiedName.ProjectId == node.Declaration.ProjectId);

            _externalRemoveCommand.Execute(param);
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
        public AnnotateDeclarationCommand AnnotateDeclarationCommand { get; set; }
        public OpenDesignerCommand OpenDesignerCommand { get; set; }
        public OpenProjectPropertiesCommand OpenProjectPropertiesCommand { get; set; }
        public SetAsStartupProjectCommand SetAsStartupProjectCommand { get; set; }
        public RenameCommand RenameCommand { get; set; }
        public CodeExplorerMoveToFolderCommand MoveToFolderCommand { get; set; }
        public IndentCommand IndenterCommand { get; set; }
        public CodeExplorerFindAllReferencesCommand FindAllReferencesCommand { get; set; }
        public CodeExplorerFindAllImplementationsCommand FindAllImplementationsCommand { get; set; }
        public CommandBase CollapseAllSubnodesCommand { get; } 
        public CopyResultsCommand CopyResultsCommand { get; set; }
        public CommandBase ExpandAllSubnodesCommand { get; }
        public ImportCommand ImportCommand { get; set; }
        public UpdateFromFilesCommand UpdateFromFilesCommand { get; set; }
        public ReplaceProjectContentsFromFilesCommand ReplaceProjectContentsFromFilesCommand { get; set; }
        public ExportCommand ExportCommand { get; set; }
        public ExportAllCommand ExportAllCommand { get; set; }
        public DeleteCommand DeleteCommand { get; set; }
        public CommandBase RemoveCommand { get; }
        public PrintCommand PrintCommand { get; set; }
        public AddRemoveReferencesCommand AddRemoveReferencesCommand { get; set; }
        public ICommand ClearSearchCommand { get; }
        public CommandBase SyncCodePaneCommand { get; }
        public CodeExplorerExtractInterfaceCommand CodeExplorerExtractInterfaceCommand { get; set; }

        public CodeExplorerMoveToFolderDragAndDropCommand MoveToFolderDragAndDropCommand { get; set; }

    public ICodeExplorerNode FindVisibleNodeForDeclaration(Declaration declaration)
        {
            if (declaration == null)
            {
                return null;
            }

            var project = Projects.OfType<CodeExplorerProjectViewModel>().FirstOrDefault(proj =>
                (proj.Declaration?.ProjectId ?? string.Empty).Equals(declaration.ProjectId));

            if (declaration.DeclarationType == DeclarationType.Project)
            {
                return project;
            }

            var child = FindChildNodeForDeclaration(project, declaration);
            return child;
        }

        private ICodeExplorerNode FindChildNodeForDeclaration(ICodeExplorerNode node, Declaration declaration)
        {
            if (node is null || declaration is null)
            {
                return null;
            }

            if (node.Declaration.DeclarationType == declaration.DeclarationType &&
                node.Declaration.QualifiedName.Equals(declaration.QualifiedName))
            {
                return node;
            }

            return node.Children.OfType<CodeExplorerItemViewModel>()
                .Select(child => FindChildNodeForDeclaration(child, declaration))
                .FirstOrDefault(child => child != null);
        }

        private void SwitchNodeState(ICodeExplorerNode node, bool expandedState)
        {
            node.IsExpanded = expandedState;

            foreach (var item in node.Children)
            {
                item.IsExpanded = expandedState;
                SwitchNodeState(item, expandedState);
            }
        }

        /// <summary>
        /// Works backward from the passed node and expands all parents to make it visible.
        /// </summary>
        /// <param name="node"></param>
        private void ExpandToNode(ICodeExplorerNode node)
        {
            while (true)
            {
                node = node.Parent;
                if (node == null)
                {
                    return;
                }
                node.IsExpanded = true;
            }
        }

        private bool CanExecuteExportAllCommand => ExportAllCommand?.CanExecute(SelectedItem) ?? false;

        public Visibility ExportVisibility => _vbe.Kind == VBEKind.Standalone || CanExecuteExportAllCommand ? Visibility.Collapsed : Visibility.Visible;

        public Visibility ExportAllVisibility => CanExecuteExportAllCommand ? Visibility.Visible : Visibility.Collapsed;

        public Visibility VB6Visibility => _vbe.Kind == VBEKind.Standalone ? Visibility.Visible : Visibility.Collapsed;

        public Visibility VBAVisibility => _vbe.Kind == VBEKind.Hosted ? Visibility.Visible : Visibility.Collapsed;

        public bool AllowDragAndDrop { get; internal set; }

        private void GeneralSettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            RefreshDragAndDropSetting();
        }

        private void RefreshDragAndDropSetting()
        {
            AllowDragAndDrop = _generalSettingsProvider.Read().EnableFolderDragAndDrop;
        }

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= HandleStateChanged;
                _state.ModuleStateChanged -= ParserState_ModuleStateChanged;
                _generalSettingsProvider.SettingsChanged -= GeneralSettingsChanged;
            }
        }
    }
}

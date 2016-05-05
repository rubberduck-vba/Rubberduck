using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using Rubberduck.UI.Refactorings;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using MessageBox = Rubberduck.UI.MessageBox;

// ReSharper disable CanBeReplacedWithTryCastAndCheckForNull

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerViewModel : ViewModelBase, IDisposable
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly NewUnitTestModuleCommand _newUnitTestModuleCommand;
        private readonly Indenter _indenter;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly FindAllReferencesCommand _findAllReferences;
        private readonly FindAllImplementationsCommand _findAllImplementations;
        private readonly SaveFileDialog _saveFileDialog;
        private readonly OpenFileDialog _openFileDialog;
        private readonly Dispatcher _dispatcher;

        public CodeExplorerViewModel(VBE vbe,
            RubberduckParserState state,
            INavigateCommand navigateCommand,
            NewUnitTestModuleCommand newUnitTestModuleCommand,
            Indenter indenter,
            ICodePaneWrapperFactory wrapperFactory,
            FindAllReferencesCommand findAllReferences,
            FindAllImplementationsCommand findAllImplementations,
            SaveFileDialog saveFileDialog,
            OpenFileDialog openFileDialog)
        {
            _dispatcher = Dispatcher.CurrentDispatcher;

            _vbe = vbe;
            _state = state;
            _newUnitTestModuleCommand = newUnitTestModuleCommand;
            _indenter = indenter;
            _wrapperFactory = wrapperFactory;
            _findAllReferences = findAllReferences;
            _findAllImplementations = findAllImplementations;
            _saveFileDialog = saveFileDialog;
            _openFileDialog = openFileDialog;
            _state.StateChanged += ParserState_StateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;

            _openFileDialog.AddExtension = true;
            _openFileDialog.AutoUpgradeEnabled = true;
            _openFileDialog.CheckFileExists = true;
            _openFileDialog.Multiselect = false;
            _openFileDialog.ShowHelp = false;   // we don't want 1996's file picker.
            _openFileDialog.Filter = @"VB Files|*.cls;*.bas;*.frm";
            _openFileDialog.CheckFileExists = true;

            _navigateCommand = navigateCommand;
            _contextMenuNavigateCommand = new DelegateCommand(ExecuteContextMenuNavigateCommand,
                CanExecuteContextMenuNavigateCommand);
            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand, _ => CanRefresh);
            _addTestModuleCommand = new DelegateCommand(ExecuteAddTestModuleCommand);
            _addStdModuleCommand = new DelegateCommand(ExecuteAddStdModuleCommand, CanAddModule);
            _addClsModuleCommand = new DelegateCommand(ExecuteAddClsModuleCommand, CanAddModule);
            _addFormCommand = new DelegateCommand(ExecuteAddFormCommand, CanAddModule);

            _openDesignerCommand = new DelegateCommand(ExecuteOpenDesignerCommand, _ => CanExecuteShowDesignerCommand);
            _indenterCommand = new DelegateCommand(ExecuteIndenterCommand, _ => CanExecuteIndenterCommand);
            _renameCommand = new DelegateCommand(ExecuteRenameCommand, _ => CanExecuteRenameCommand);
            _findAllReferencesCommand = new DelegateCommand(ExecuteFindAllReferencesCommand, _ => CanExecuteFindAllReferencesCommand);
            _findAllImplementationsCommand = new DelegateCommand(ExecuteFindAllImplementationsCommand, _ => CanExecuteFindAllImplementationsCommand);

            _printCommand = new DelegateCommand(ExecutePrintCommand, CanExecutePrintCommand);
            _importCommand = new DelegateCommand(ExecuteImportCommand, CanExecuteImportCommand);
            _exportCommand = new DelegateCommand(ExecuteExportCommand, CanExecuteExportCommand);
            _removeCommand = new DelegateCommand(ExecuteRemoveCommand, CanExecuteRemoveCommand);
        }

        private readonly ICommand _refreshCommand;

        public ICommand RefreshCommand
        {
            get { return _refreshCommand; }
        }

        private readonly ICommand _addTestModuleCommand;

        public ICommand AddTestModuleCommand
        {
            get { return _addTestModuleCommand; }
        }

        private readonly ICommand _addStdModuleCommand;

        public ICommand AddStdModuleCommand
        {
            get { return _addStdModuleCommand; }
        }

        private readonly ICommand _addClsModuleCommand;

        public ICommand AddClsModuleCommand
        {
            get { return _addClsModuleCommand; }
        }

        private readonly ICommand _addFormCommand;

        public ICommand AddFormCommand
        {
            get { return _addFormCommand; }
        }

        private readonly ICommand _openDesignerCommand;

        public ICommand OpenDesignerCommand
        {
            get { return _openDesignerCommand; }
        }

        private readonly ICommand _indenterCommand;

        public ICommand IndenterCommand
        {
            get { return _indenterCommand; }
        }

        private readonly ICommand _renameCommand;

        public ICommand RenameCommand
        {
            get { return _renameCommand; }
        }

        private readonly ICommand _findAllReferencesCommand;

        public ICommand FindAllReferencesCommand
        {
            get { return _findAllReferencesCommand; }
        }

        private readonly ICommand _findAllImplementationsCommand;

        public ICommand FindAllImplementationsCommand
        {
            get { return _findAllImplementationsCommand; }
        }

        private readonly INavigateCommand _navigateCommand;

        public ICommand NavigateCommand
        {
            get { return _navigateCommand; }
        }

        private readonly ICommand _printCommand;

        public ICommand PrintCommand
        {
            get { return _printCommand; }
        }

        private readonly ICommand _importCommand;

        public ICommand ImportCommand
        {
            get { return _importCommand; }
        }

        private readonly ICommand _exportCommand;

        public ICommand ExportCommand
        {
            get { return _exportCommand; }
        }

        private readonly ICommand _removeCommand;

        public ICommand RemoveCommand
        {
            get { return _removeCommand; }
        }

        private readonly ICommand _contextMenuNavigateCommand;

        public ICommand ContextMenuNavigateCommand
        {
            get { return _contextMenuNavigateCommand; }
        }

        public string Description
        {
            get
            {
                if (SelectedItem is CodeExplorerProjectViewModel)
                {
                    return ((CodeExplorerProjectViewModel) SelectedItem).Declaration.DescriptionString;
                }

                if (SelectedItem is CodeExplorerComponentViewModel)
                {
                    return ((CodeExplorerComponentViewModel) SelectedItem).Declaration.DescriptionString;
                }

                if (SelectedItem is CodeExplorerMemberViewModel)
                {
                    return ((CodeExplorerMemberViewModel) SelectedItem).Declaration.DescriptionString;
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
                OnPropertyChanged("CanExecuteShowDesignerCommand");
                OnPropertyChanged("CanExecutePrintCommand");
                OnPropertyChanged("CanExecuteExportCommand");
                OnPropertyChanged("CanExecuteRemoveCommand");
                OnPropertyChanged("PanelTitle");
                OnPropertyChanged("Description");
                // ReSharper restore ExplicitCallerInfoArgument
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

        private bool CanAddModule(object param)
        {
            return _vbe.ActiveVBProject != null;
        }

        private bool CanExecuteContextMenuNavigateCommand(object param)
        {
            return SelectedItem != null && SelectedItem.QualifiedSelection.HasValue;
        }

        private bool CanExecutePrintCommand(object param)
        {
            return SelectedItem is CodeExplorerComponentViewModel;
        }

        private bool CanExecuteImportCommand(object param)
        {
            return SelectedItem is CodeExplorerProjectViewModel || 
                   SelectedItem is CodeExplorerComponentViewModel ||
                   SelectedItem is CodeExplorerMemberViewModel;
        }

        private bool CanExecuteExportCommand(object param)
        {
            if (!(SelectedItem is CodeExplorerComponentViewModel))
            {
                return false;
            }

            var node = (CodeExplorerComponentViewModel) SelectedItem;
            var componentType = node.Declaration.QualifiedName.QualifiedModuleName.Component.Type;
            return _exportableFileExtensions.Select(s => s.Key).Contains(componentType);
        }

        private bool CanExecuteRemoveCommand(object param)
        {
            return SelectedItem is CodeExplorerComponentViewModel;
        }

        private bool CanExecuteShowDesignerCommand
        {
            get
            {
                var declaration = GetSelectedDeclaration();
                return declaration != null && declaration.DeclarationType == DeclarationType.ClassModule &&
                       declaration.QualifiedName.QualifiedModuleName.Component.Designer != null;
            }
        }

        public bool CanExecuteIndenterCommand
        {
            get
            {
                return _state.Status == ParserState.Ready && !(SelectedItem is CodeExplorerCustomFolderViewModel) &&
                       !(SelectedItem is CodeExplorerErrorNodeViewModel);
            }
        }

        public bool CanExecuteRenameCommand
        {
            get
            {
                return _state.Status == ParserState.Ready && !(SelectedItem is CodeExplorerCustomFolderViewModel) &&
                       !(SelectedItem is CodeExplorerErrorNodeViewModel);
            }
        }

        public bool CanExecuteFindAllReferencesCommand
        {
            get
            {
                return _state.Status == ParserState.Ready && !(SelectedItem is CodeExplorerCustomFolderViewModel) &&
                       !(SelectedItem is CodeExplorerErrorNodeViewModel);
            }
        }

        private bool CanExecuteFindAllImplementationsCommand
        {
            get
            {
                return _state.Status == ParserState.Ready &&
                       (SelectedItem is CodeExplorerComponentViewModel ||
                        SelectedItem is CodeExplorerMemberViewModel);
            }
        }

        private ObservableCollection<CodeExplorerItemViewModel> _projects;

        public ObservableCollection<CodeExplorerItemViewModel> Projects
        {
            get { return _projects; }
            set
            {
                _projects = value;
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

            if (
                userDeclarations.Any(
                    grouping => grouping.All(declaration => declaration.DeclarationType != DeclarationType.Project)))
            {
                return;
            }

            var newProjects = new ObservableCollection<CodeExplorerItemViewModel>(userDeclarations.Select(grouping =>
                new CodeExplorerProjectViewModel(
                    grouping.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Project),
                    grouping)));

            UpdateNodes(Projects, newProjects);
            Projects = newProjects;
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

        private delegate void AddErrorNode(
            CodeExplorerItemViewModel projectNode, CodeExplorerItemViewModel folderNode, string componentName);

        private void AddComponentErrorNode(CodeExplorerItemViewModel projectNode, CodeExplorerItemViewModel folderNode,
            string componentName)
        {
            Projects.Remove(projectNode);
            RemoveFailingComponent(projectNode, componentName);

            folderNode.AddChild(new CodeExplorerErrorNodeViewModel(componentName));
            Projects.Add(projectNode);
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

        private void ExecuteRefreshCommand(object param)
        {
            _state.OnParseRequested(this);
        }

        private void ExecuteAddTestModuleCommand(object param)
        {
            _newUnitTestModuleCommand.NewUnitTestModule();
        }

        private void ExecuteAddStdModuleCommand(object param)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
        }

        private void ExecuteAddClsModuleCommand(object param)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
        }

        private void ExecuteAddFormCommand(object param)
        {
            _vbe.ActiveVBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_MSForm);
        }

        private void ExecuteOpenDesignerCommand(object param)
        {
            GetSelectedDeclaration().QualifiedName.QualifiedModuleName.Component.DesignerWindow().Visible = true;
        }

        private void ExecuteIndenterCommand(object param)
        {
            if (!SelectedItem.QualifiedSelection.HasValue)
            {
                return;
            }

            if (SelectedItem is CodeExplorerProjectViewModel)
            {
                _indenter.Indent(SelectedItem.QualifiedSelection.Value.QualifiedName.Project);
            }

            if (SelectedItem is CodeExplorerComponentViewModel)
            {
                _indenter.Indent(SelectedItem.QualifiedSelection.Value.QualifiedName.Component);
            }

            if (SelectedItem is CodeExplorerMemberViewModel)
            {
                var arg = new NavigateCodeEventArgs(SelectedItem.QualifiedSelection.Value);
                NavigateCommand.Execute(arg);

                _indenter.IndentCurrentProcedure();
            }
        }

        private void ExecuteRenameCommand(object obj)
        {
            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(_vbe, view, _state, new MessageBox(), _wrapperFactory);
                var refactoring = new RenameRefactoring(_vbe, factory, new MessageBox(), _state);

                refactoring.Refactor(GetSelectedDeclaration());
            }
        }

        private void ExecuteFindAllReferencesCommand(object obj)
        {
            _findAllReferences.Execute(GetSelectedDeclaration());
        }

        private void ExecuteFindAllImplementationsCommand(object obj)
        {
            _findAllImplementations.Execute(GetSelectedDeclaration());
        }

        private void ExecuteContextMenuNavigateCommand(object obj)
        {
            // ReSharper disable once PossibleInvalidOperationException
            // CanExecute protects against this
            var arg = new NavigateCodeEventArgs(SelectedItem.QualifiedSelection.Value);

            NavigateCommand.Execute(arg);
        }

        private void ExecutePrintCommand(object obj)
        {
            var node = (CodeExplorerComponentViewModel) SelectedItem;
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck",
                component.Name + ".txt");

            var text = component.CodeModule.Lines[1, component.CodeModule.CountOfLines];

            var printDoc = new PrintDocument {DocumentName = path};
            var pd = new PrintDialog
            {
                Document = printDoc,
                AllowSelection = true,
                AllowSomePages = true
            };

            if (pd.ShowDialog() == DialogResult.OK)
            {
                printDoc.PrintPage += (sender, printPageArgs) =>
                {
                    var font = new Font(new FontFamily("Consolas"), 10, FontStyle.Regular);
                    printPageArgs.Graphics.DrawString(text, font, Brushes.Black, 0, 0, new StringFormat());
                };
                printDoc.Print();
            }
        }

        private void ExecuteImportCommand(object param)
        {
            // I know this will never be null because of the CanExecute
            var project = GetSelectedDeclaration().QualifiedName.QualifiedModuleName.Project;

            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var fileExt = '.' + _openFileDialog.FileName.Split('.').Last();
                if (!_exportableFileExtensions.Select(s => s.Value).Contains(fileExt))
                {
                    return;
                }

                project.VBComponents.Import(_openFileDialog.FileName);
            }
        }

        private readonly Dictionary<vbext_ComponentType, string> _exportableFileExtensions = new Dictionary<vbext_ComponentType, string>
        {
            { vbext_ComponentType.vbext_ct_StdModule, ".bas" },
            { vbext_ComponentType.vbext_ct_ClassModule, ".cls" },
            { vbext_ComponentType.vbext_ct_Document, ".cls" },
            { vbext_ComponentType.vbext_ct_MSForm, ".frm" }
        };

        private void ExecuteExportCommand(object param)
        {
            var node = (CodeExplorerComponentViewModel)SelectedItem;
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            string ext;
            _exportableFileExtensions.TryGetValue(component.Type, out ext);

            _saveFileDialog.FileName = component.Name + ext;
            if (_saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                component.Export(_saveFileDialog.FileName);
            }
        }

        private void ExecuteRemoveCommand(object param)
        {
        }

        private Declaration GetSelectedDeclaration()
        {
            if (SelectedItem is CodeExplorerProjectViewModel)
            {
                return ((CodeExplorerProjectViewModel) SelectedItem).Declaration;
            }

            if (SelectedItem is CodeExplorerComponentViewModel)
            {
                return ((CodeExplorerComponentViewModel) SelectedItem).Declaration;
            }

            if (SelectedItem is CodeExplorerMemberViewModel)
            {
                return ((CodeExplorerMemberViewModel) SelectedItem).Declaration;
            }

            return null;
        }

        public void Dispose()
        {
            _saveFileDialog.Dispose();
            _openFileDialog.Dispose();
        }
    }
}

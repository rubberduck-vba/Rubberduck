using System;
using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Interaction.Navigation;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that locates all references to a specified identifier, or of the active code module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllReferencesCommand : CommandBase, IDisposable
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly IVBE _vbe;
        private readonly IUiDispatcher _uiDispatcher;

        public FindAllReferencesCommand(INavigateCommand navigateCommand, IMessageBox messageBox,
            RubberduckParserState state, IVBE vbe, ISearchResultsWindowViewModel viewModel,
            SearchResultPresenterInstanceManager presenterService, IUiDispatcher uiDispatcher)
             : base(LogManager.GetCurrentClassLogger())
        {
            _navigateCommand = navigateCommand;
            _messageBox = messageBox;
            _state = state;
            _vbe = vbe;
            _viewModel = viewModel;
            _presenterService = presenterService;
            _uiDispatcher = uiDispatcher;

            _state.StateChanged += _state_StateChanged;
        }

        private Declaration FindNewDeclaration(Declaration declaration)
        {
            return _state.DeclarationFinder
                .MatchName(declaration.IdentifierName)
                .SingleOrDefault(d => d.ProjectId == declaration.ProjectId
                    && d.ComponentName == declaration.ComponentName
                    && d.ParentScope == declaration.ParentScope
                    && d.DeclarationType == declaration.DeclarationType);            
        }

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready) { return; }

            if (_viewModel == null) { return; }

            _uiDispatcher.InvokeAsync(UpdateTab);
        }

        private void UpdateTab()
        {
            try
            {
                var findReferenceTabs = _viewModel.Tabs.Where(
                    t => t.Header.StartsWith(RubberduckUI.AllReferences_Caption.Replace("'{0}'", ""))).ToList();

                foreach (var tab in findReferenceTabs)
                {
                    var newTarget = FindNewDeclaration(tab.Target);
                    if (newTarget == null)
                    {
                        tab.CloseCommand.Execute(null);
                        return;
                    }

                    var vm = CreateViewModel(newTarget);
                    if (vm.SearchResults.Any())
                    {
                        tab.SearchResults = vm.SearchResults;
                        tab.Target = vm.Target;
                    }
                    else
                    {
                        tab.CloseCommand.Execute(null);
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Exception thrown while trying to update the find all references tab.");
            }
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                using (var selectedComponent = _vbe.SelectedVBComponent)
                {
                    if ((activePane == null || activePane.IsWrappingNullReference)
                        && !(selectedComponent?.HasDesigner ?? false))
                    {
                        return false;
                    }
                }
            }

            var target = FindTarget(parameter);
            var canExecute = target != null;

            return canExecute;
        }

        protected override void OnExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            var declaration = FindTarget(parameter);
            if (declaration == null)
            {
                return;
            }

            var viewModel = CreateViewModel(declaration);
            if (!viewModel.SearchResults.Any())
            {
                _messageBox.NotifyWarn(string.Format(RubberduckUI.AllReferences_NoneFound, declaration.IdentifierName), RubberduckUI.Rubberduck);
                return;
            }

            if (viewModel.SearchResults.Count == 1)
            {
                _navigateCommand.Execute(viewModel.SearchResults.Single().GetNavigationArgs());
                return;
            }

            _viewModel.AddTab(viewModel);
            _viewModel.SelectedTab = viewModel;

            try
            {
                var presenter = _presenterService.Presenter(_viewModel);
                presenter.Show();
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }
        }

        private SearchResultsViewModel CreateViewModel(Declaration declaration)
        {
            var results = declaration.References.Distinct().Select(reference =>
                new SearchResultItem(
                    reference.ParentNonScoping,
                    new NavigateCodeEventArgs(reference.QualifiedModuleName, reference.Selection),
                    GetModuleLine(reference.QualifiedModuleName, reference.Selection.StartLine)));
            
            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllReferencesTabFormat, declaration.IdentifierName), declaration, results);

            return viewModel;
        }

        private string GetModuleLine(QualifiedModuleName module, int line)
        {
            var component = _state.ProjectsProvider.Component(module);
            using (var codeModule = component.CodeModule)
            {
                return codeModule.GetLines(line, 1).Trim();
            }
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            using (var activePane = _vbe.ActiveCodePane)
            {
                bool findDesigner;
                using (var selectedComponent = _vbe.SelectedVBComponent)
                {
                    findDesigner = activePane != null && !activePane.IsWrappingNullReference
                                                      && (selectedComponent?.HasDesigner ?? false);
                }

                return findDesigner
                    ? FindFormDesignerTarget()
                    : FindCodePaneTarget(activePane);
            }
        }

        private Declaration FindCodePaneTarget(ICodePane codePane)
        {
            return _state.FindSelectedDeclaration(codePane);
        }

        private Declaration FindFormDesignerTarget(QualifiedModuleName? qualifiedModuleName = null)
        {
            if (qualifiedModuleName.HasValue)
            {
                return FindFormDesignerTarget(qualifiedModuleName.Value);
            }

            string projectId;
            using (var activeProject = _vbe.ActiveVBProject)
            {
                projectId = activeProject.ProjectId;
            }
            var component = _vbe.SelectedVBComponent;

            if (component?.HasDesigner ?? false)
            {
                DeclarationType selectedType;
                string selectedName;
                using (var selectedControls = component.SelectedControls)
                {
                    var selectedCount = selectedControls.Count;
                    if (selectedCount > 1)
                    {
                        return null;
                    }

                    // Cannot use DeclarationType.UserForm, parser only assigns UserForms the ClassModule flag
                    (selectedType, selectedName) = selectedCount == 0
                        ? (DeclarationType.ClassModule, component.Name)
                        : (DeclarationType.Control, selectedControls[0].Name);
                }
                return _state.DeclarationFinder
                    .MatchName(selectedName)
                    .SingleOrDefault(m => m.ProjectId == projectId
                        && m.DeclarationType.HasFlag(selectedType)
                        && m.ComponentName == component.Name);                
            }
            return null;
        }

        private Declaration FindFormDesignerTarget(QualifiedModuleName qualifiedModuleName)
        {
            var projectId = qualifiedModuleName.ProjectId;
            var component = _state.ProjectsProvider.Component(qualifiedModuleName);

            if (component?.HasDesigner ?? false)
            {
                return _state.DeclarationFinder
                    .MatchName(qualifiedModuleName.Name)
                    .SingleOrDefault(m => m.ProjectId == projectId
                                          && m.DeclarationType.HasFlag(qualifiedModuleName.ComponentType)
                                          && m.ComponentName == component.Name);
            }

            return null;
        }


        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }
        }
    }
}

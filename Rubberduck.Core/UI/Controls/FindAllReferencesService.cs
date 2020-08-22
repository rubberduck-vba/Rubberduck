using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Controls
{
    public class FindAllReferencesService : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly INavigateCommand _navigateCommand;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly IUiDispatcher _uiDispatcher;

        public FindAllReferencesService(
            INavigateCommand navigateCommand, 
            IMessageBox messageBox,
            RubberduckParserState state, 
            ISearchResultsWindowViewModel viewModel,
            SearchResultPresenterInstanceManager presenterService, 
            IUiDispatcher uiDispatcher)
        {
            _navigateCommand = navigateCommand;
            _messageBox = messageBox;
            _state = state;
            _viewModel = viewModel;
            _presenterService = presenterService;
            _uiDispatcher = uiDispatcher;

            _state.StateChanged += _state_StateChanged;
        }

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready || _viewModel == null)
            {
                return;
            }

            _uiDispatcher.InvokeAsync(UpdateTab);
        }

        public void FindAllReferences(Declaration declaration)
        {
            if (_state.Status != ParserState.Ready)
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
                _logger.Error(e);
            }
        }

        public void FindAllReferences(Declaration declaration, ReferenceInfo reference)
        {
            if (_state.Status != ParserState.Ready ||
                !(declaration is ProjectDeclaration project))
            {
                return;
            }

            var usages = _state.DeclarationFinder.FindAllReferenceUsesInProject(project, reference, out var referenceProject);
            if (!usages.Any())
            {
                _messageBox.NotifyWarn(string.Format(RubberduckUI.AllReferences_NoneFoundReference, referenceProject.IdentifierName), RubberduckUI.Rubberduck);
                return;
            }

            if (usages.Count > 1000 &&
                !_messageBox.ConfirmYesNo(string.Format(RubberduckUI.AllReferences_PerformanceWarning, referenceProject.IdentifierName, usages.Count),
                    RubberduckUI.PerformanceWarningCaption))
            {
                return;
            }

            var viewModel = CreateViewModel(project, referenceProject, usages);
            _viewModel.AddTab(viewModel);
            _viewModel.SelectedTab = viewModel;

            try
            {
                var presenter = _presenterService.Presenter(_viewModel);
                presenter.Show();
            }
            catch (Exception e)
            {
                _logger.Error(e);
            }
        }

        private string GetModuleLine(QualifiedModuleName module, int line)
        {
            var component = _state.ProjectsProvider.Component(module);
            using (var codeModule = component.CodeModule)
            {
                return codeModule.GetLines(line, 1).Trim();
            }
        }

        private SearchResultsViewModel CreateViewModel(ProjectDeclaration project, ProjectDeclaration reference, IEnumerable<IdentifierReference> usages)
        {
            var results = usages.Select(usage =>
                new SearchResultItem(usage.ParentNonScoping,
                    new NavigateCodeEventArgs(usage.QualifiedModuleName, usage.Selection),
                    GetModuleLine(usage.QualifiedModuleName, usage.Selection.StartLine)));

            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllReferencesTabFormat, reference.IdentifierName), project, results);

            return viewModel;
        }

        private SearchResultsViewModel CreateViewModel(Declaration declaration)
        {
            var results = declaration.References
                .Where(reference => !reference.IsArrayAccess)
                .Distinct()
                .Select(reference =>
                    new SearchResultItem(
                        reference.ParentNonScoping,
                        new NavigateCodeEventArgs(reference.QualifiedModuleName, reference.Selection),
                        GetModuleLine(reference.QualifiedModuleName, reference.Selection.StartLine)));

            var accessor = declaration.DeclarationType.HasFlag(DeclarationType.PropertyGet) ? "(get)"
                : declaration.DeclarationType.HasFlag(DeclarationType.PropertyLet) ? "(let)"
                : declaration.DeclarationType.HasFlag(DeclarationType.PropertySet) ? "(set)"
                : string.Empty;

            var tabCaption = $"{declaration.IdentifierName} {accessor}".Trim();


            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllReferencesTabFormat, tabCaption), declaration, results);

            return viewModel;
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
                _logger.Error(exception, "Exception thrown while trying to update the find all references tab.");
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }
            _isDisposed = true;
        }
    }
}

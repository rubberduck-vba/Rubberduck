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
    public class FindAllReferencesAction : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly INavigateCommand _navigateCommand;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly IUiDispatcher _uiDispatcher;

        public FindAllReferencesAction(
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

        public void FindAllReferences(Declaration target)
        {
            if (_state.Status != ParserState.Ready)
            {
                _logger.Debug($"ParserState is {_state.Status}. This action requires a Ready state.");
                return;
            }

            var viewModel = CreateViewModel(target);
            if (!Confirm(target.IdentifierName, viewModel.SearchResults.Count))
            {
                return;
            }

            ShowResults(viewModel);
        }

        public void FindAllReferences(ProjectDeclaration project, ReferenceInfo reference)
        {
            if (_state.Status != ParserState.Ready)
            {
                _logger.Debug($"ParserState is {_state.Status}. This action requires a Ready state.");
                return;
            }

            var usages = _state.DeclarationFinder.FindAllReferenceUsesInProject(project, reference, out var referenceProject).ToList();
            if (referenceProject == null)
            {
                return;
            }
            if (!Confirm(referenceProject.IdentifierName, usages.Count))
            {
                return;
            }

            var viewModel = CreateViewModel(project, referenceProject.IdentifierName, usages);
            ShowResults(viewModel);
        }

        private void ShowResults(SearchResultsViewModel viewModel)
        {
            if (viewModel.SearchResults.Count == 1)
            {
                viewModel.NavigateCommand.Execute(viewModel.SearchResults[0].GetNavigationArgs());
                return;
            }

            try
            {
                _viewModel.AddTab(viewModel);
                _viewModel.SelectedTab = viewModel;

                var presenter = _presenterService.Presenter(_viewModel);
                presenter.Show();
            }
            catch (Exception e)
            {
                _logger.Error(e);
            }
        }

        private bool Confirm(string identifier, int referencesFound)
        {
            const int threshold = 1000;
            if (referencesFound == 0)
            {
                _messageBox.NotifyWarn(
                    string.Format(RubberduckUI.AllReferences_NoneFoundReference, identifier), 
                    RubberduckUI.Rubberduck);
                return false;
            }

            if (referencesFound > threshold)
            {
                return _messageBox.ConfirmYesNo(
                    string.Format(RubberduckUI.AllReferences_PerformanceWarning, identifier, referencesFound),
                    RubberduckUI.PerformanceWarningCaption);
            }

            return true;
        }


        private SearchResultsViewModel CreateViewModel(Declaration declaration, string identifier = null, IEnumerable<IdentifierReference> references = null)
        {
            var nameRefs = (references ?? declaration.References)
                .Where(reference => !reference.IsArrayAccess)
                .Distinct()
                .GroupBy(reference => reference.QualifiedModuleName)
                .ToDictionary(group => group.Key);

            var argRefs = (declaration is ParameterDeclaration parameter
                    ? parameter.ArgumentReferences
                    : Enumerable.Empty<ArgumentReference>())
                .Distinct()
                .GroupBy(argRef => argRef.QualifiedModuleName)
                .ToDictionary(group => group.Key);

            var results = new List<SearchResultItem>();
            var modules = nameRefs.Keys.Concat(argRefs.Keys).Distinct();
            foreach (var qualifiedModuleName in modules)
            {
                var component = _state.ProjectsProvider.Component(qualifiedModuleName);
                if (component == null)
                {
                    _logger.Warn($"Could not retrieve the IVBComponent for module '{qualifiedModuleName}'.");
                    continue;
                }
                var module = component.CodeModule;

                if (nameRefs.TryGetValue(qualifiedModuleName, out var identifierReferences))
                {
                    foreach (var identifierReference in identifierReferences)
                    {
                        var (context, selection) = identifierReference.HighlightSelection(module);
                        var result = new SearchResultItem(
                            identifierReference.ParentNonScoping,
                            new NavigateCodeEventArgs(qualifiedModuleName, identifierReference.Selection),
                            context, selection);
                        results.Add(result);
                    }
                }

                if (argRefs.TryGetValue(qualifiedModuleName, out var argReferences))
                {
                    foreach (var argumentReference in argReferences)
                    {
                        var (context, selection) = argumentReference.HighlightSelection(module);
                        var result = new SearchResultItem(
                            argumentReference.ParentNonScoping,
                            new NavigateCodeEventArgs(qualifiedModuleName, argumentReference.Selection),
                            context, selection);
                        results.Add(result);
                    }
                }
            }

            var accessor = declaration.DeclarationType.HasFlag(DeclarationType.PropertyGet) ? "(get)"
                : declaration.DeclarationType.HasFlag(DeclarationType.PropertyLet) ? "(let)"
                : declaration.DeclarationType.HasFlag(DeclarationType.PropertySet) ? "(set)"
                : string.Empty;

            var tabCaption = $"{identifier ?? declaration.IdentifierName} {accessor}".Trim();


            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllReferencesTabFormat, tabCaption), declaration, 
                results.OrderBy(item => item.ParentScope.QualifiedModuleName.ToString())
                       .ThenBy(item => item.Selection)
                       .ToList());

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

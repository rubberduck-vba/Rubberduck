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
    public class FindAllImplementationsService : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly INavigateCommand _navigateCommand;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly IUiDispatcher _uiDispatcher;

        public FindAllImplementationsService(INavigateCommand navigateCommand, IMessageBox messageBox,
            RubberduckParserState state, ISearchResultsWindowViewModel viewModel,
            SearchResultPresenterInstanceManager presenterService, IUiDispatcher uiDispatcher)
        {
            _navigateCommand = navigateCommand;
            _messageBox = messageBox;
            _state = state;
            _viewModel = viewModel;
            _presenterService = presenterService;
            _uiDispatcher = uiDispatcher;

            _state.StateChanged += _state_StateChanged;
        }

        public bool CanFind(Declaration declaration)
        {
            return declaration is ModuleBodyElementDeclaration moduleBody &&
                   moduleBody.Accessibility == Accessibility.Public &&
                   declaration.ParentDeclaration is ClassModuleDeclaration;
        }

        public void FindAllImplementations(Declaration declaration)
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
                Console.WriteLine(e);
            }
        }

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready || _viewModel == null)
            {
                return;
            }

            _uiDispatcher.InvokeAsync(UpdateTab);
        }

        private void UpdateTab()
        {
            try
            {
                var findImplementationsTabs = _viewModel.Tabs.Where(
                    t => t.Header.StartsWith(RubberduckUI.AllImplementations_Caption.Replace("'{0}'", ""))).ToList();

                foreach (var tab in findImplementationsTabs)
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
                _logger.Error(exception, "Exception thrown while trying to update the find implementations tab.");
            }
        }

        private Declaration FindNewDeclaration(Declaration declaration)
        {
            return _state.AllUserDeclarations.SingleOrDefault(item =>
                item.ProjectId == declaration.ProjectId &&
                item.ComponentName == declaration.ComponentName &&
                item.ParentScope == declaration.ParentScope &&
                item.IdentifierName == declaration.IdentifierName &&
                item.DeclarationType == declaration.DeclarationType);
        }

        private SearchResultsViewModel CreateViewModel(Declaration target)
        {
            IEnumerable<Declaration> implementations;
            if (target is ClassModuleDeclaration classModule)
            {
                implementations = _state.DeclarationFinder.FindAllImplementationsOfInterface(classModule);
            }
            else if (target is IInterfaceExposable member && member.IsInterfaceMember)
            {
                implementations = _state.DeclarationFinder.FindInterfaceImplementationMembers(target);
            }
            else
            {
                implementations = target is ModuleBodyElementDeclaration implementation
                    ? _state.DeclarationFinder.FindInterfaceImplementationMembers(implementation.InterfaceMemberImplemented)
                    : Enumerable.Empty<Declaration>();
            }

            var results = implementations.Select(declaration =>
                new SearchResultItem(
                    declaration.ParentScopeDeclaration,
                    new NavigateCodeEventArgs(declaration.QualifiedName.QualifiedModuleName, declaration.Selection),
                    GetModuleLine(declaration.QualifiedName.QualifiedModuleName, declaration.Selection.StartLine)));

            var accessor = target.DeclarationType.HasFlag(DeclarationType.PropertyGet) ? "(get)"
                : target.DeclarationType.HasFlag(DeclarationType.PropertyLet) ? "(let)"
                : target.DeclarationType.HasFlag(DeclarationType.PropertySet) ? "(set)"
                : string.Empty;

            var tabCaption = $"{target.IdentifierName} {accessor}".Trim();

            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllImplementationsTabFormat, tabCaption), target, results);

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

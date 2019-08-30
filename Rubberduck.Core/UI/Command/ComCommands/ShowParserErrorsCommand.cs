using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    public interface IShowParserErrorsCommand : ICommand, IDisposable { }

    [ComVisible(false)]
    public class ShowParserErrorsCommand : ComCommandBase, IShowParserErrorsCommand
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly IUiDispatcher _uiDispatcher;

        public ShowParserErrorsCommand(
            INavigateCommand navigateCommand, 
            RubberduckParserState state, 
            ISearchResultsWindowViewModel viewModel, 
            SearchResultPresenterInstanceManager presenterService,
            IUiDispatcher uiDispatcher, 
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _navigateCommand = navigateCommand;
            _state = state;
            _viewModel = viewModel;
            _presenterService = presenterService;
            _uiDispatcher = uiDispatcher;

            _state.StateChanged += _state_StateChanged;
        }

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (_viewModel == null) { return; }

            if (_state.Status != ParserState.Error && _state.Status != ParserState.Parsed) { return; }

            _uiDispatcher.InvokeAsync(UpdateTab);
        }

        private void UpdateTab()
        {
            try
            {
                if (_viewModel == null)
                {
                    return;
                }

                var vm = CreateViewModel();

                var tab = _viewModel.Tabs.FirstOrDefault(t => t.Header == RubberduckUI.Parser_ParserError);
                if (tab != null)
                {
                    if (_state.Status != ParserState.Error)
                    {
                        tab.CloseCommand.Execute(null);
                    }
                    else
                    {
                        tab.SearchResults = vm.SearchResults;
                    }
                }
                else if (_state.Status == ParserState.Error)
                {
                    _viewModel.AddTab(vm);
                    _viewModel.SelectedTab = vm;
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Exception thrown while trying to update the parser errors tab.");
            }
        }

        protected override void OnExecute(object parameter)
        {
            if (_viewModel == null)
            {
                return;
            }

            try
            {
                UpdateTab();
                var presenter = _presenterService.Presenter(_viewModel);
                presenter.Show();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private SearchResultsViewModel CreateViewModel()
        {
            var errors = from error in _state.ModuleExceptions
                let declaration = FindModuleDeclaration(error.Item1)
                where declaration != null
                select new SearchResultItem(declaration, error.Item2.GetNavigateCodeEventArgs(declaration), error.Item2.Message);

            var searchResultItems = errors as IList<SearchResultItem> ?? errors.ToList();
            var viewModel = new SearchResultsViewModel(_navigateCommand,RubberduckUI.Parser_ParserError, null, searchResultItems);
            return viewModel;
        }

        private Declaration FindModuleDeclaration(QualifiedModuleName module)
        {
            var projectId = module.ProjectId;
            var project = _state.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                    .SingleOrDefault(item => item.ProjectId == projectId);

            var result = _state.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .SingleOrDefault(item => item.QualifiedName.QualifiedModuleName.Equals(module));

            // FIXME dirty hack for project.Scope in case project is null. Clean up!
            var declaration = new Declaration(new QualifiedMemberName(module, module.ComponentName), project, project?.Scope, module.ComponentName, null, false, false, Accessibility.Global, DeclarationType.ProceduralModule, false, null, true);
            return result ?? declaration; // module isn't in parser state - give it a dummy declaration, just so the ViewModel has something to chew on
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

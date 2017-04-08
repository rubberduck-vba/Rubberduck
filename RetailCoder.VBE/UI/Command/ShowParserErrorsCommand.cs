using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Input;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    public interface IShowParserErrorsCommand : ICommand, IDisposable { }

    [ComVisible(false)]
    public class ShowParserErrorsCommand : CommandBase, IShowParserErrorsCommand
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;

        public ShowParserErrorsCommand(INavigateCommand navigateCommand,
            RubberduckParserState state,
            ISearchResultsWindowViewModel viewModel,
            SearchResultPresenterInstanceManager presenterService)
            : base(LogManager.GetCurrentClassLogger())
        {
            _navigateCommand = navigateCommand;
            _state = state;
            _viewModel = viewModel;
            _presenterService = presenterService;

            _state.StateChanged += _state_StateChanged;
        }

        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (_viewModel == null) { return; }

            if (_state.Status != ParserState.Error && _state.Status != ParserState.Parsed) { return; }
            
            UiDispatcher.InvokeAsync(UpdateTab);
        }

        private void UpdateTab()
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

        protected override void ExecuteImpl(object parameter)
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

        private Declaration FindModuleDeclaration(IVBComponent component)
        {
            var components = component.Collection;
            var refProject = components.Parent;
            {
                var projectId = refProject.HelpFile;
                var project = _state.AllUserDeclarations.SingleOrDefault(item =>
                    item.DeclarationType == DeclarationType.Project && item.ProjectId == projectId);

                var result = _state.AllUserDeclarations.SingleOrDefault(item => 
                    item.ProjectId == component.Collection.Parent.HelpFile
                    && item.QualifiedName.QualifiedModuleName.ComponentName == component.Name
                    && (item.DeclarationType == DeclarationType.ClassModule || item.DeclarationType == DeclarationType.ProceduralModule));

                // FIXME dirty hack for project.Scope in case project is null. Clean up!
                var declaration = new Declaration(new QualifiedMemberName(new QualifiedModuleName(component), component.Name), project, project?.Scope, component.Name, null, false, false, Accessibility.Global, DeclarationType.ProceduralModule, false, null, true);
                return result ?? declaration; // module isn't in parser state - give it a dummy declaration, just so the ViewModel has something to chew on
            }
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

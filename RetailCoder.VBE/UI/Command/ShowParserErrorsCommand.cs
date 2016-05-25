using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command
{
    public interface IShowParserErrorsCommand : ICommand { }

    [ComVisible(false)]
    public class ShowParserErrorsCommand : CommandBase, IShowParserErrorsCommand
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;

        public ShowParserErrorsCommand(INavigateCommand navigateCommand, RubberduckParserState state, ISearchResultsWindowViewModel viewModel, SearchResultPresenterInstanceManager presenterService)
        {
            _navigateCommand = navigateCommand;
            _state = state;
            _viewModel = viewModel;
            _presenterService = presenterService;
        }

        public override void Execute(object parameter)
        {
            if (_state.Status != ParserState.Error)
            {
                return;
            }

            var viewModel = CreateViewModel();
            if (_viewModel == null)
            {
                return;
            }

            var oldTab = _viewModel.SelectedTab;
            
            _viewModel.AddTab(viewModel);
            _viewModel.SelectedTab = viewModel;

            if (oldTab != null)
            {
                oldTab.CloseCommand.Execute(null);
            }

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

        private Declaration FindModuleDeclaration(VBComponent component)
        {
            var projectId = component.Collection.Parent.HelpFile;

            var project = _state.AllUserDeclarations.SingleOrDefault(item => 
                item.DeclarationType == DeclarationType.Project && item.ProjectId == projectId);

            var result = _state.AllUserDeclarations.SingleOrDefault(item => item.ProjectId == component.Collection.Parent.HelpFile
                                                             && item.QualifiedName.QualifiedModuleName.ComponentName == component.Name
                                                             && (item.DeclarationType == DeclarationType.ClassModule || item.DeclarationType == DeclarationType.ProceduralModule));
           
            // FIXME dirty hack for project.Scope in case project is null. Clean up!
            var declaration = new Declaration(new QualifiedMemberName(new QualifiedModuleName(component), component.Name), project, project == null ? null : project.Scope, component.Name, null, false, false, Accessibility.Global, DeclarationType.ProceduralModule, false, null, false);
            return result ?? declaration; // module isn't in parser state - give it a dummy declaration, just so the ViewModel has something to chew on
        }
    }
}

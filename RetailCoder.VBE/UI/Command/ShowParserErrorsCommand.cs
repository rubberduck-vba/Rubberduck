using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;

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

        private SearchResultsViewModel CreateViewModel()
        {
            var errors = from error in _state.ModuleExceptions
                let declaration = FindModuleDeclaration(error.Item1)
                select new SearchResultItem(declaration, error.Item2.GetNavigateCodeEventArgs(declaration), error.Item2.Message);

            var viewModel = new SearchResultsViewModel(_navigateCommand, "Parser Errors", null, errors.ToList());
            return viewModel;
        }

        private Declaration FindModuleDeclaration(VBComponent component)
        {
            return _state.AllUserDeclarations.Single(item => item.Project == component.Collection.Parent
                                                             && item.QualifiedName.QualifiedModuleName.Component == component 
                                                             && (item.DeclarationType == DeclarationType.Class || item.DeclarationType == DeclarationType.Module));
        }
    }
}
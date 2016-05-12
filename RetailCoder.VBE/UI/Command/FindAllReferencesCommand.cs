using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that locates all references to a specified identifier, or of the active code module.
    /// </summary>
    [ComVisible(false)]
    public class FindAllReferencesCommand : CommandBase
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;
        private readonly VBE _vbe;

        public FindAllReferencesCommand(INavigateCommand navigateCommand, IMessageBox messageBox, RubberduckParserState state, VBE vbe, ISearchResultsWindowViewModel viewModel, SearchResultPresenterInstanceManager presenterService)
        {
            _navigateCommand = navigateCommand;
            _messageBox = messageBox;
            _state = state;
            _vbe = vbe;
            _viewModel = viewModel;
            _presenterService = presenterService;
        }

        public override bool CanExecute(object parameter)
        {
            if (_vbe.ActiveCodePane == null && _state.Status != ParserState.Ready)
            {
                return false;
            }

            // todo: make this work for Code/Project Explorer context menus too (may require a new command implementation)
            var target = FindTarget(parameter);
            var canExecute = target != null;

            return canExecute;
        }

        public override void Execute(object parameter)
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
                _messageBox.Show(string.Format(RubberduckUI.AllReferences_NoneFound, declaration.IdentifierName), RubberduckUI.Rubberduck, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private SearchResultsViewModel CreateViewModel(Declaration declaration)
        {
            var results = declaration.References.Select(reference =>
                new SearchResultItem(
                    reference.ParentNonScoping,
                    new NavigateCodeEventArgs(reference.QualifiedModuleName, reference.Selection), 
                    reference.QualifiedModuleName.Component.CodeModule.Lines[reference.Selection.StartLine, 1].Trim()));
            
            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllReferencesTabFormat, declaration.IdentifierName), declaration, results);

            return viewModel;
        }

        private Declaration FindTarget(object parameter)
        {
            var declaration = parameter as Declaration;
            if (declaration != null)
            {
                return declaration;
            }

            return _state.FindSelectedDeclaration(_vbe.ActiveCodePane);
        }
    }
}
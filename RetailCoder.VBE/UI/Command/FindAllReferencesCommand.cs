using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;

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
        private readonly IActiveCodePaneEditor _editor;
        private readonly ISearchResultsWindowViewModel _viewModel;
        private readonly SearchResultPresenterInstanceManager _presenterService;

        public FindAllReferencesCommand(INavigateCommand navigateCommand, IMessageBox messageBox, RubberduckParserState state, IActiveCodePaneEditor editor, ISearchResultsWindowViewModel viewModel, SearchResultPresenterInstanceManager presenterService)
        {
            _navigateCommand = navigateCommand;
            _messageBox = messageBox;
            _state = state;
            _editor = editor;
            _viewModel = viewModel;
            _presenterService = presenterService;
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
                    reference.QualifiedModuleName.QualifyMemberName(reference.ParentScope.Split('.').Last()),
                    reference.Selection,
                    _editor.GetLines(reference.Selection)));
            
            var viewModel = new SearchResultsViewModel(_navigateCommand,
                string.Format(RubberduckUI.SearchResults_AllReferencesTabFormat, declaration.IdentifierName), results);

            return viewModel;
        }

        private Declaration FindTarget(object parameter)
        {
            var declaration = parameter as Declaration;
            if (declaration == null)
            {
                var selection = _editor.GetSelection();
                if (selection != null)
                {
                    declaration = _state.AllUserDeclarations
                        .SingleOrDefault(item => item.QualifiedName.QualifiedModuleName == selection.Value.QualifiedName 
                            && (item.QualifiedSelection.Selection.ContainsFirstCharacter(selection.Value.Selection)
                                || 
                                item.References.Any(reference => reference.Selection.ContainsFirstCharacter(selection.Value.Selection))));
                }

                if (declaration == null)
                {
                    return null;
                }
            }
            return declaration;
        }
    }
}
using System;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_RenameCommand : CommandBase, IDisposable
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IRenameDialog _view;
        private readonly IMessageBox _msgBox;

        public CodeExplorer_RenameCommand(VBE vbe, RubberduckParserState state, IRenameDialog view, IMessageBox msgBox)
        {
            _vbe = vbe;
            _state = state;
            _view = view;
            _msgBox = msgBox;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready && !(parameter is CodeExplorerCustomFolderViewModel) &&
                   !(parameter is CodeExplorerErrorNodeViewModel);
        }

        public override void Execute(object parameter)
        {
            var factory = new RenamePresenterFactory(_vbe, _view, _state, _msgBox);
            var refactoring = new RenameRefactoring(_vbe, factory, _msgBox, _state);

            refactoring.Refactor(GetSelectedDeclaration((CodeExplorerItemViewModel)parameter));
        }

        private Declaration GetSelectedDeclaration(CodeExplorerItemViewModel node)
        {
            if (node is CodeExplorerProjectViewModel)
            {
                return ((CodeExplorerProjectViewModel)node).Declaration;
            }

            if (node is CodeExplorerComponentViewModel)
            {
                return ((CodeExplorerComponentViewModel)node).Declaration;
            }

            if (node is CodeExplorerMemberViewModel)
            {
                return ((CodeExplorerMemberViewModel)node).Declaration;
            }

            return null;
        }

        public void Dispose()
        {
            if (_view != null)
            {
                _view.Dispose();
            }
        }
    }
}
using System;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_RenameCommand : CommandBase, IDisposable
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly RenameDialog _view;

        public CodeExplorer_RenameCommand(VBE vbe, RubberduckParserState state, ICodePaneWrapperFactory wrapperFactory, RenameDialog view)
        {
            _vbe = vbe;
            _state = state;
            _wrapperFactory = wrapperFactory;
            _view = view;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready && !(parameter is CodeExplorerCustomFolderViewModel) &&
                   !(parameter is CodeExplorerErrorNodeViewModel);
        }

        public override void Execute(object parameter)
        {
            var factory = new RenamePresenterFactory(_vbe, _view, _state, new MessageBox(), _wrapperFactory);
            var refactoring = new RenameRefactoring(_vbe, factory, new MessageBox(), _state);

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
            if (_view != null && !_view.IsDisposed)
            {
                _view.Dispose();
            }
        }
    }
}
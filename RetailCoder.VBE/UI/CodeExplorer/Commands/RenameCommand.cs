using System;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class RenameCommand : CommandBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly RenameDialog _view;
        private readonly IMessageBox _msgBox;

        public RenameCommand(IVBE vbe, RenameDialog view, RubberduckParserState state, IMessageBox msgBox) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _view = view;
            _msgBox = msgBox;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Ready && parameter is ICodeExplorerDeclarationViewModel;
        }

        protected override void ExecuteImpl(object parameter)
        {
            var factory = new RenamePresenterFactory(_vbe, _view, _state, _msgBox);
            var refactoring = new RenameRefactoring(_vbe, factory, _msgBox, _state);

            refactoring.Refactor(((ICodeExplorerDeclarationViewModel)parameter).Declaration);
        }

        public void Dispose()
        {
            _view?.Dispose();
        }
    }
}

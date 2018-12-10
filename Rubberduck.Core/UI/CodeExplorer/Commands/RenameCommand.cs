using System;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class RenameCommand : CommandBase, IDisposable
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IRefactoringDialog<RenameViewModel> _view;
        private readonly IMessageBox _msgBox;

        public RenameCommand(IVBE vbe, IRefactoringDialog<RenameViewModel> view, RubberduckParserState state, IMessageBox msgBox, IRewritingManager rewritingManager) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _rewritingManager = rewritingManager;
            _view = view;
            _msgBox = msgBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready && parameter is ICodeExplorerDeclarationViewModel;
        }

        protected override void OnExecute(object parameter)
        {
            var factory = new RenamePresenterFactory(_vbe, _view, _state);
            var refactoring = new RenameRefactoring(_vbe, factory, _msgBox, _state, _state.ProjectsProvider, _rewritingManager);

            refactoring.Refactor(((ICodeExplorerDeclarationViewModel)parameter).Declaration);
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

            _view?.Dispose();
            _isDisposed = true;
        }
    }
}

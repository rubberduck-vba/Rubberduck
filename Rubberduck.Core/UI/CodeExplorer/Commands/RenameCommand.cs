using System;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public sealed class RenameCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly IMessageBox _msgBox;

        public RenameCommand(IVBE vbe, RubberduckParserState state,
            IMessageBox msgBox, IRefactoringPresenterFactory factory) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _factory = factory;
            _msgBox = msgBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready && parameter is ICodeExplorerDeclarationViewModel;
        }

        protected override void OnExecute(object parameter)
        {
            var refactoring = new RenameRefactoring(_vbe, _factory, _msgBox, _state);
            refactoring.Refactor(((ICodeExplorerDeclarationViewModel)parameter).Declaration);
        }
    }
}

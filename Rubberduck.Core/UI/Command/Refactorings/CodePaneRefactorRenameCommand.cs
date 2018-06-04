using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory<IRenamePresenter> _factory;

        public CodePaneRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory<IRenamePresenter> factory) 
            : base (vbe)
        {
            _state = state;
            _messageBox = messageBox;
            _factory = factory;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            Declaration target;
            using (var activePane = Vbe.ActiveCodePane)
            {
                if (Vbe.ActiveCodePane == null || activePane.IsWrappingNullReference)
                {
                    return false;
                }
            
                target = _state.FindSelectedDeclaration(activePane);
            }

            return _state.Status == ParserState.Ready 
                && target != null 
                && target.IsUserDefined 
                && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            Declaration target;
            using (var activePane = Vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return;
                }

                if (parameter != null)
                {
                    target = parameter as Declaration;
                }
                else
                {
                    target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane);
                }
            }

            if (target == null || !target.IsUserDefined)
            {
                return;
            }

            var refactoring = new RenameRefactoring(Vbe, _factory, _messageBox, _state);
            refactoring.Refactor(target);
        }
    }
}

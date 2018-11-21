using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceParameterCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;

        public RefactorIntroduceParameterCommand (IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRewritingManager rewritingManager)
            :base(vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            var selection = Vbe.GetActiveSelection();

            if (!selection.HasValue)
            {
                return false;
            }

            var target = _state.AllUserDeclarations.FindVariable(selection.Value);

            return target != null
                && !_state.IsNewOrModified(target.QualifiedModuleName)
                && target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
        }

        protected override void OnExecute(object parameter)
        {
            var selection = Vbe.GetActiveSelection();

            if (!selection.HasValue)
            {
                return;
            }

            var refactoring = new IntroduceParameterRefactoring(Vbe, _state, _messageBox, _rewritingManager);
            refactoring.Refactor(selection.Value);
        }
    }
}

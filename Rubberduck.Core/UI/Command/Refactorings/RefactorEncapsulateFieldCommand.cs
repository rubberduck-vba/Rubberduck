using System.Runtime.InteropServices;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorEncapsulateFieldCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly Indenter _indenter;
        private readonly IRefactoringPresenterFactory _factory;

        public RefactorEncapsulateFieldCommand(IVBE vbe, RubberduckParserState state, Indenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager)
            : base(vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _indenter = indenter;
            _factory = factory;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            Declaration target;
            using (var pane = Vbe.ActiveCodePane)
            {
                if (pane == null || _state.Status != ParserState.Ready)
                {
                    return false;
                }

                target = _state.FindSelectedDeclaration(pane);
            }
            return target != null
                && target.DeclarationType == DeclarationType.Variable
                && !target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member)
                && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        protected override void OnExecute(object parameter)
        {
            using(var activePane = Vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return;
                }
            }

            var refactoring = new EncapsulateFieldRefactoring(_state, Vbe, _indenter, _factory, _rewritingManager);
            refactoring.Refactor();
        }
    }
}

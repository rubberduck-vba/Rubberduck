using System.Runtime.InteropServices;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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

        public RefactorEncapsulateFieldCommand(IVBE vbe, RubberduckParserState state, Indenter indenter, IRewritingManager rewritingManager)
            : base(vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _indenter = indenter;
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

            using (var view = new EncapsulateFieldDialog(new EncapsulateFieldViewModel(_state, _indenter)))
            {
                var factory = new EncapsulateFieldPresenterFactory(Vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(Vbe, _indenter, factory, _rewritingManager);
                refactoring.Refactor();
            }
        }
    }
}

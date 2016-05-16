using System.Diagnostics;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorEncapsulateFieldCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorEncapsulateFieldCommand(VBE vbe, RubberduckParserState state)
            : base(vbe)
        {
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            if (pane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var target = _state.FindSelectedDeclaration(pane);

            var canExecute = target != null 
                && target.DeclarationType == DeclarationType.Variable
                && !target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);

            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            using (var view = new EncapsulateFieldDialog())
            {
                var factory = new EncapsulateFieldPresenterFactory(Vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(Vbe, factory);
                refactoring.Refactor();
            }
        }
    }
}

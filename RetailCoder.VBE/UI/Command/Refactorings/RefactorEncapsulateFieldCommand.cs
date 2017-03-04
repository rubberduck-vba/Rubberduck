using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.UI.Refactorings;
using Rubberduck.SmartIndenter;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorEncapsulateFieldCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly Indenter _indenter;

        public RefactorEncapsulateFieldCommand(IVBE vbe, RubberduckParserState state, Indenter indenter)
            : base(vbe)
        {
            _state = state;
            _indenter = indenter;
        }

        protected override bool CanExecuteImpl(object parameter)
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

        protected override void ExecuteImpl(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            using (var view = new EncapsulateFieldDialog(_state, _indenter))
            {
                var factory = new EncapsulateFieldPresenterFactory(Vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(Vbe, _indenter, factory);
                refactoring.Refactor();
            }
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.RefactorEncapsulateField; }
        }
    }
}

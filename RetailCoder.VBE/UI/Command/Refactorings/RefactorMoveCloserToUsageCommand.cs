using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorMoveCloserToUsageCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgbox;

        public RefactorMoveCloserToUsageCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox)
            :base(vbe)
        {
            _state = state;
            _msgbox = msgbox;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.RefactorMoveCloserToUsage; }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane);
            return target != null 
                && (target.DeclarationType == DeclarationType.Variable || target.DeclarationType == DeclarationType.Constant)
                && target.References.Any();
        }

        protected override void ExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            var module = pane.CodeModule;
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }

                var selection = new QualifiedSelection(new QualifiedModuleName(module.Parent), pane.Selection);

                var refactoring = new MoveCloserToUsageRefactoring(Vbe, _state, _msgbox);
                refactoring.Refactor(selection);
            }
        }
    }
}

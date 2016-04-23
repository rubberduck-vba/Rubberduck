using System.Diagnostics;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorMoveCloserToUsageCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorMoveCloserToUsageCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory)
            : base(vbe, editor)
        {
            _state = state;
            _wrapperWrapperFactory = wrapperWrapperFactory;
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane);
            var canExecute = target != null 
                && (target.DeclarationType == DeclarationType.Variable || target.DeclarationType == DeclarationType.Constant)
                && target.References.Any();

            Debug.WriteLine("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }
            var codePane = _wrapperWrapperFactory.Create(Vbe.ActiveCodePane);
            var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), codePane.Selection);

            var refactoring = new MoveCloserToUsageRefactoring(_state, Editor, new MessageBox());
            refactoring.Refactor(selection);
        }
    }
}
using System.Diagnostics;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorIntroduceParameterCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorIntroduceParameterCommand (VBE vbe, RubberduckParserState state, ICodePaneWrapperFactory wrapperWrapperFactory)
            :base(vbe)
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

            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            var target = _state.AllUserDeclarations.FindTarget(selection, new []{DeclarationType.Variable, DeclarationType.Constant});

            var canExecute = target != null && target.ParentScopeDeclaration != null && target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);

            Debug.WriteLine("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            var refactoring = new IntroduceParameterRefactoring(Vbe, _state, new MessageBox());
            refactoring.Refactor(selection);
        }
    }
}
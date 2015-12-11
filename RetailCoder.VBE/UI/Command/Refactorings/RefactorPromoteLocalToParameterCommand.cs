using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.PromoteLocalToParameter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorPromoteLocalToParameterCommand : RefactorCommandBase
    {
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorPromoteLocalToParameterCommand (VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory)
            : base(vbe, editor)
        {
            _wrapperWrapperFactory = wrapperWrapperFactory;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }
            var codePane = _wrapperWrapperFactory.Create(Vbe.ActiveCodePane);
            var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), codePane.Selection);

            var refactoring = new PromoteLocalToParameterRefactoring();
            refactoring.Refactor(selection);
        }
    }
}
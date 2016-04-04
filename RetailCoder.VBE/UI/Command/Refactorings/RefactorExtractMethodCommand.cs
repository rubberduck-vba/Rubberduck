using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorExtractMethodCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor)
            : base (vbe, editor)
        {
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return false;
            }

            var selection = Vbe.ActiveCodePane.GetSelection();
            var target = _state.AllDeclarations.FindSelectedDeclaration(selection, DeclarationExtensions.ProcedureTypes, d => ((ParserRuleContext)d.Context.Parent).GetSelection());
            return _state.Status == ParserState.Ready && target != null;
        }

        public override void Execute(object parameter)
        {
            var factory = new ExtractMethodPresenterFactory(Editor, _state.AllDeclarations);
            var refactoring = new ExtractMethodRefactoring(factory, Editor);
            refactoring.InvalidSelection += HandleInvalidSelection;
            refactoring.Refactor();
        }
    }
}
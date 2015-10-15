using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.UI.ParserProgress;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        public RefactorExtractMethodCommand(VBE vbe, ParsingProgressPresenter parserProgress, IActiveCodePaneEditor editor)
            : base (vbe, parserProgress, editor)
        {
        }

        public override void Execute(object parameter)
        {
            var result = ParserProgress.Parse(Vbe.ActiveVBProject);

            var declarations = result.Declarations;
            var factory = new ExtractMethodPresenterFactory(Editor, declarations);
            var refactoring = new ExtractMethodRefactoring(factory, Editor);
            refactoring.InvalidSelection += HandleInvalidSelection;
            refactoring.Refactor();
        }
    }
}
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.UI.ParserProgress;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        public RefactorExtractMethodCommand(VBE vbe, IParsingProgressPresenter parserProgress, IActiveCodePaneEditor editor)
            : base (vbe, parserProgress, editor)
        {
        }

        public override void Execute(object parameter)
        {
            var result = ParserProgress.Parse(Vbe.ActiveVBProject);

            var factory = new ExtractMethodPresenterFactory(Editor, result);
            var refactoring = new ExtractMethodRefactoring(factory, Editor);
            refactoring.InvalidSelection += HandleInvalidSelection;
            refactoring.Refactor();
        }
    }
}
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractMethodCommand : RefactorCommandBase
    {
        public RefactorExtractMethodCommand (VBE ide, IRubberduckParser parser, IActiveCodePaneEditor editor)
            : base (ide, parser, editor)
        {
        }

        public override void Execute(object parameter)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, _ide.ActiveVBProject);

            var declarations = result.Declarations;
            var factory = new ExtractMethodPresenterFactory(_editor, declarations);
            var refactoring = new ExtractMethodRefactoring(factory, _editor);
            refactoring.InvalidSelection += refactoring_InvalidSelection;
            refactoring.Refactor();
        }
    }
}
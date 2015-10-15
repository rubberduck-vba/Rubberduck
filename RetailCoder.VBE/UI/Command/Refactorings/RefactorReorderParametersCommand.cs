using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.ParserProgress;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    class RefactorReorderParametersCommand : RefactorCommandBase
    {
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorReorderParametersCommand(VBE vbe, IParsingProgressPresenter parserProgress, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe, parserProgress, editor)
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
            // duplicates ReorderParameters Implementation until here... extract common method?
            // TryGetQualifiedSelection?
            var result = ParserProgress.Parse(Vbe.ActiveVBProject);

            using (var view = new ReorderParametersDialog())
            {
                var factory = new ReorderParametersPresenterFactory(Editor, view, result, new MessageBox());
                var refactoring = new ReorderParametersRefactoring(factory, Editor, new MessageBox());
                refactoring.Refactor(selection);
            }
        }
    }
}

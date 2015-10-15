using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.ParserProgress;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRemoveParametersCommand : RefactorCommandBase
    {
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorRemoveParametersCommand(VBE vbe, IRubberduckParser parser, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe, parser, editor)
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
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(Vbe.ActiveVBProject);

            using (var view = new RemoveParametersDialog())
            {
                var factory = new RemoveParametersPresenterFactory(Editor, view, result, new MessageBox());
                var refactoring = new RemoveParametersRefactoring(factory, Editor);
                refactoring.Refactor(selection);
            }
        }
    }
}
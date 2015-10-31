using System.Linq;
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
        private readonly IRubberduckParser _parser;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorRemoveParametersCommand(VBE vbe, IRubberduckParser parser, IParsingProgressPresenter parserProgress, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe, parserProgress, editor)
        {
            _parser = parser;
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

            using (var view = new RemoveParametersDialog())
            {
                var factory = new RemoveParametersPresenterFactory(Editor, view, _parser.State, new MessageBox());
                var refactoring = new RemoveParametersRefactoring(factory, Editor);
                refactoring.Refactor(selection);
            }
        }
    }
}
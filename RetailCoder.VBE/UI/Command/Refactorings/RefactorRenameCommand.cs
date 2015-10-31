using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.ParserProgress;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRenameCommand : RefactorCommandBase
    {
        private readonly IRubberduckParser _parser;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorRenameCommand(VBE vbe, IRubberduckParser parser, IParsingProgressPresenter parserProgress, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory) 
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

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _parser.State, new MessageBox(), _wrapperWrapperFactory);
                var refactoring = new RenameRefactoring(factory, Editor, new MessageBox());

                var target = parameter as Declaration;
                if (target == null)
                {
                    refactoring.Refactor();
                }
                else
                {
                    refactoring.Refactor(target);
                }
            }
        }
    }
}
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.ParserProgress;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRenameCommand : RefactorCommandBase
    {
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorRenameCommand(VBE vbe, IRubberduckParser parser, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory) 
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

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, result, new MessageBox(), _wrapperWrapperFactory);
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
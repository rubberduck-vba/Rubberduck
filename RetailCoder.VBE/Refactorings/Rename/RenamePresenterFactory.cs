using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenterFactory : IRefactoringPresenterFactory<RenamePresenter>
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringDialog<RenameViewModel> _view;
        private readonly RubberduckParserState _state;

        public RenamePresenterFactory(IVBE vbe, IRefactoringDialog<RenameViewModel> view, RubberduckParserState state)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
        }

        public RenamePresenter Create()
        {
            QualifiedSelection qualifiedSelection;
            using (var codePane = _vbe.ActiveCodePane)
            {
                using (var codeModule = codePane.CodeModule)
                {
                    qualifiedSelection = codePane.IsWrappingNullReference
                        ? new QualifiedSelection()
                        : new QualifiedSelection(new QualifiedModuleName(codeModule.Parent), codePane.Selection);
                }
            }
            return new RenamePresenter(_view, new RenameModel(_vbe, _state, qualifiedSelection));
        }
    }
}

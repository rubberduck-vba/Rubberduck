using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenterFactory : IRefactoringPresenterFactory<RenamePresenter>
    {
        private readonly IVBE _vbe;
        private readonly IRenameDialog _view;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RenamePresenterFactory(IVBE vbe, IRenameDialog view, RubberduckParserState state, IMessageBox messageBox)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
            _messageBox = messageBox;
        }

        public RenamePresenter Create()
        {
            var codePane = _vbe.ActiveCodePane;
            var qualifiedSelection = codePane.IsWrappingNullReference
                ? new QualifiedSelection()
                : new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), codePane.Selection);

            return new RenamePresenter(_view, new RenameModel(_vbe, _state, qualifiedSelection, _messageBox));
        }
    }
}

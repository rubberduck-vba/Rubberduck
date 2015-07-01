using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenterFactory : IRefactoringPresenterFactory<RenamePresenter>
    {
        private readonly IRenameView _view;
        private readonly RenameModel _model;

        public RenamePresenterFactory(VBE vbe, IRenameView view, VBProjectParseResult parseResult)
        {
            _view = view;

            _model = new RenameModel(vbe, parseResult, vbe.ActiveCodePane.GetSelection());

        }

        public RenamePresenter Create()
        {
            return new RenamePresenter(_view, _model);
        }
    }
}
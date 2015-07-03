using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenterFactory : IRefactoringPresenterFactory<RenamePresenter>
    {
        private readonly VBE _vbe;
        private readonly IRenameView _view;
        private readonly VBProjectParseResult _parseResult;

        public RenamePresenterFactory(VBE vbe, IRenameView view, VBProjectParseResult parseResult)
        {
            _vbe = vbe;
            _view = view;
            _parseResult = parseResult;
        }

        public RenamePresenter Create()
        {
            return new RenamePresenter(_view, new RenameModel(_vbe, _parseResult, _vbe.ActiveCodePane.GetSelection()));
        }
    }
}
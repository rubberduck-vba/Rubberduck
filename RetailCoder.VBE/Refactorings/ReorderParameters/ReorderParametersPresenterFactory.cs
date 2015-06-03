using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenterFactory : IRefactoringPresenterFactory<ReorderParametersPresenter>
    {
        private readonly VBE _vbe;
        private readonly IReorderParametersDialog _view;
        private readonly VBProjectParseResult _parseResult;

        public ReorderParametersPresenterFactory(VBE vbe, IReorderParametersDialog view,
            VBProjectParseResult parseResult)
        {
            _vbe = vbe;
            _view = view;
            _parseResult = parseResult;
        }

        public ReorderParametersPresenter Create()
        {
            var selection = _vbe.ActiveCodePane.GetSelection();
            var model = new ReorderParametersModel(_parseResult, selection);

            return new ReorderParametersPresenter(_view, model);
        }
    }
}

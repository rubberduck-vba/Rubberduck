using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenterFactory : IRefactoringPresenterFactory<ReorderParametersPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IReorderParametersView _view;
        private readonly VBProjectParseResult _parseResult;

        public ReorderParametersPresenterFactory(IActiveCodePaneEditor editor, IReorderParametersView view,
            VBProjectParseResult parseResult)
        {
            _editor = editor;
            _view = view;
            _parseResult = parseResult;
        }

        public ReorderParametersPresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new ReorderParametersModel(_parseResult, selection.Value);
            return new ReorderParametersPresenter(_view, model);
        }
    }
}

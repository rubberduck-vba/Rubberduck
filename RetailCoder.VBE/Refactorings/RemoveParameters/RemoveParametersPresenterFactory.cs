using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenterFactory : IRefactoringPresenterFactory<RemoveParametersPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IRemoveParametersView _view;
        private readonly VBProjectParseResult _parseResult;

        public RemoveParametersPresenterFactory(IActiveCodePaneEditor editor, IRemoveParametersView view,
            VBProjectParseResult parseResult)
        {
            _editor = editor;
            _view = view;
            _parseResult = parseResult;
        }

        public RemoveParametersPresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new RemoveParametersModel(_parseResult, selection.Value);
            return new RemoveParametersPresenter(_view, model);
        }
    }
}

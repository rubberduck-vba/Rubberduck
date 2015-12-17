using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenterFactory : IRefactoringPresenterFactory<RemoveParametersPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IRemoveParametersView _view;
        private readonly IRubberduckParserState _parseResult;
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenterFactory(IActiveCodePaneEditor editor, IRemoveParametersView view,
            IRubberduckParserState parseResult, IMessageBox messageBox)
        {
            _editor = editor;
            _view = view;
            _parseResult = parseResult;
            _messageBox = messageBox;
        }

        public RemoveParametersPresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new RemoveParametersModel(_parseResult, selection.Value, _messageBox);
            return new RemoveParametersPresenter(_view, model, _messageBox);
        }
    }
}

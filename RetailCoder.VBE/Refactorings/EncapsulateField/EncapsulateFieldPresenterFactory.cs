using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPresenterFactory : IRefactoringPresenterFactory<EncapsulateFieldPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IEncapsulateFieldView _view;
        private readonly RubberduckParserState _parseResult;
        private readonly IMessageBox _messageBox;

        public EncapsulateFieldPresenterFactory(RubberduckParserState parseResult, IActiveCodePaneEditor editor, 
            IEncapsulateFieldView view, IMessageBox messageBox)
        {
            _editor = editor;
            _view = view;
            _parseResult = parseResult;
            _messageBox = messageBox;
        }

        public EncapsulateFieldPresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new EncapsulateFieldModel(_parseResult, selection.Value, _messageBox);
            return new EncapsulateFieldPresenter(_view, model, _messageBox);
        }
    }
}
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPresenterFactory : IRefactoringPresenterFactory<EncapsulateFieldPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IEncapsulateFieldView _view;
        private readonly RubberduckParserState _parseResult;

        public EncapsulateFieldPresenterFactory(RubberduckParserState parseResult, IActiveCodePaneEditor editor, IEncapsulateFieldView view)
        {
            _editor = editor;
            _view = view;
            _parseResult = parseResult;
        }

        public EncapsulateFieldPresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new EncapsulateFieldModel(_parseResult, selection.Value);
            return new EncapsulateFieldPresenter(_view, model);
        }
    }
}
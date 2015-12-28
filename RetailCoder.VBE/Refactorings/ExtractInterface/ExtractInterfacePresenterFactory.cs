using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfacePresenterFactory : IRefactoringPresenterFactory<ExtractInterfacePresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IExtractInterfaceView _view;
        private readonly RubberduckParserState _parseResult;

        public ExtractInterfacePresenterFactory(RubberduckParserState parseResult, IActiveCodePaneEditor editor, IExtractInterfaceView view)
        {
            _editor = editor;
            _view = view;
            _parseResult = parseResult;
        }

        public ExtractInterfacePresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new ExtractInterfaceModel(_parseResult, selection.Value);
            return new ExtractInterfacePresenter(_view, model);
        }
    }
}
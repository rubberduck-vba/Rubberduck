using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfacePresenterFactory : IRefactoringPresenterFactory<ExtractInterfacePresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IExtractInterfaceView _view;
        private readonly RubberduckParserState _state;

        public ExtractInterfacePresenterFactory(RubberduckParserState state, IActiveCodePaneEditor editor, IExtractInterfaceView view)
        {
            _editor = editor;
            _view = view;
            _state = state;
        }

        public ExtractInterfacePresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new ExtractInterfaceModel(_state, selection.Value);
            if (!model.Members.Any())
            {
                // don't show the UI if there's no member to extract
                return null;
            }

            return new ExtractInterfacePresenter(_view, model);
        }
    }
}
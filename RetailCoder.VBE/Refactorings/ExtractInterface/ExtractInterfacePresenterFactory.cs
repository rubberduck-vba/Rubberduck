using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfacePresenterFactory : IRefactoringPresenterFactory<ExtractInterfacePresenter>
    {
        private readonly IVBE _vbe;
        private readonly IExtractInterfaceDialog _view;
        private readonly RubberduckParserState _state;

        public ExtractInterfacePresenterFactory(IVBE vbe, RubberduckParserState state, IExtractInterfaceDialog view)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
        }

        public ExtractInterfacePresenter Create()
        {
            var pane = _vbe.ActiveCodePane;
            if (pane == null || pane.IsWrappingNullReference)
            {
                return null;
            }
            var selection = pane.GetQualifiedSelection();
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

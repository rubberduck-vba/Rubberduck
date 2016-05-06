using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfacePresenterFactory : IRefactoringPresenterFactory<ExtractInterfacePresenter>
    {
        private readonly VBE _vbe;
        private readonly IExtractInterfaceDialog _view;
        private readonly RubberduckParserState _state;

        public ExtractInterfacePresenterFactory(VBE vbe, RubberduckParserState state, IExtractInterfaceDialog view)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
        }

        public ExtractInterfacePresenter Create()
        {
            var selection = _vbe.ActiveCodePane.CodeModule.GetSelection();
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
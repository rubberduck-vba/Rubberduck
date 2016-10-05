using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPresenterFactory : IRefactoringPresenterFactory<EncapsulateFieldPresenter>
    {
        private readonly VBE _vbe;
        private readonly IEncapsulateFieldDialog _view;
        private readonly RubberduckParserState _state;

        public EncapsulateFieldPresenterFactory(VBE vbe, RubberduckParserState state, IEncapsulateFieldDialog view)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
        }

        public EncapsulateFieldPresenter Create()
        {
            var selection = _vbe.ActiveCodePane.GetQualifiedSelection();
            if (!selection.HasValue)
            {
                return null;
            }

            var model = new EncapsulateFieldModel(_state, selection.Value);
            return new EncapsulateFieldPresenter(_view, model);
        }
    }
}

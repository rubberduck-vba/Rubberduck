using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPresenterFactory : IRefactoringPresenterFactory<EncapsulateFieldPresenter>
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringDialog<EncapsulateFieldViewModel> _view;
        private readonly RubberduckParserState _state;

        public EncapsulateFieldPresenterFactory(IVBE vbe, RubberduckParserState state, IRefactoringDialog<EncapsulateFieldViewModel> view)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
        }

        public EncapsulateFieldPresenter Create()
        {

            var selection = _vbe.GetActiveSelection();

            if (!selection.HasValue)
            {
                return null;
            }

            var model = new EncapsulateFieldModel(_state, selection.Value);
            return new EncapsulateFieldPresenter(_view, model);
        }
    }
}

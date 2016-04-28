using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPresenterFactory : IRefactoringPresenterFactory<EncapsulateFieldPresenter>
    {
        private readonly VBE _vbe;
        private readonly IEncapsulateFieldView _view;
        private readonly RubberduckParserState _parseResult;

        public EncapsulateFieldPresenterFactory(VBE vbe, RubberduckParserState parseResult, IEncapsulateFieldView view)
        {
            _vbe = vbe;
            _view = view;
            _parseResult = parseResult;
        }

        public EncapsulateFieldPresenter Create()
        {
            var selection = _vbe.ActiveCodePane.CodeModule.GetSelection();
            if (selection == null)
            {
                return null;
            }

            var model = new EncapsulateFieldModel(_parseResult, selection.Value);
            return new EncapsulateFieldPresenter(_view, model);
        }
    }
}
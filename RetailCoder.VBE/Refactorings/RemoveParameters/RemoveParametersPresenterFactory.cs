using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenterFactory : IRefactoringPresenterFactory<RemoveParametersPresenter>
    {
        private readonly VBE _vbe;
        private readonly IRemoveParametersView _view;
        private readonly RubberduckParserState _parseResult;
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenterFactory(VBE vbe, IRemoveParametersView view,
            RubberduckParserState parseResult, IMessageBox messageBox)
        {
            _vbe = vbe;
            _view = view;
            _parseResult = parseResult;
            _messageBox = messageBox;
        }

        public RemoveParametersPresenter Create()
        {
            var selection = _vbe.ActiveCodePane.GetQualifiedSelection();

            if (!selection.HasValue)
            {
                return null;
            }

            var model = new RemoveParametersModel(_parseResult, selection.Value, _messageBox);
            return new RemoveParametersPresenter(_view, model, _messageBox);
        }
    }
}

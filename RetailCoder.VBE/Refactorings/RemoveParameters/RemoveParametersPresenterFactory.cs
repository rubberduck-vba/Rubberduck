using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenterFactory : IRefactoringPresenterFactory<RemoveParametersPresenter>
    {
        private readonly VBE _vbe;
        private readonly IRemoveParametersDialog _view;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenterFactory(VBE vbe, IRemoveParametersDialog view,
            RubberduckParserState state, IMessageBox messageBox)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
            _messageBox = messageBox;
        }

        public RemoveParametersPresenter Create()
        {
            var selection = _vbe.ActiveCodePane.GetQualifiedSelection();

            if (!selection.HasValue)
            {
                return null;
            }

            var model = new RemoveParametersModel(_state, selection.Value, _messageBox);
            return new RemoveParametersPresenter(_view, model, _messageBox);
        }
    }
}

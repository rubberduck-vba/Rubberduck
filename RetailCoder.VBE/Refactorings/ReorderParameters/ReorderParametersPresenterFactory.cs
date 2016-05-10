using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenterFactory : IRefactoringPresenterFactory<IReorderParametersPresenter>
    {
        private readonly VBE _vbe;
        private readonly IReorderParametersDialog _view;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public ReorderParametersPresenterFactory(VBE vbe, IReorderParametersDialog view,
            RubberduckParserState state, IMessageBox messageBox)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
            _messageBox = messageBox;
        }

        public IReorderParametersPresenter Create()
        {
            var selection = _vbe.ActiveCodePane.GetQualifiedSelection();

            if (!selection.HasValue)
            {
                return null;
            }

            var model = new ReorderParametersModel(_state, selection.Value, _messageBox);
            return new ReorderParametersPresenter(_view, model, _messageBox);
        }
    }
}

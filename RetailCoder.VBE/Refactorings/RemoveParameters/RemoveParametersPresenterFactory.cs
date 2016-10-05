using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;
using Rubberduck.VBEditor.Extensions;

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
            var pane = _vbe.ActiveCodePane;
            {
                var selection = pane.GetQualifiedSelection();

                if (!selection.HasValue)
                {
                    return null;
                }

                var model = new RemoveParametersModel(_state, selection.Value, _messageBox);
                return new RemoveParametersPresenter(_view, model, _messageBox);
            }
        }
    }
}

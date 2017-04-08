using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.RemoveParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenterFactory : IRefactoringPresenterFactory<RemoveParametersPresenter>
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringDialog<RemoveParametersViewModel> _view;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenterFactory(IVBE vbe, IRefactoringDialog<RemoveParametersViewModel> view,
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
            if (pane == null || pane.IsWrappingNullReference)
            {
                return null;
            }

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

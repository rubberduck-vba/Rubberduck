using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.ReorderParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenterFactory : IRefactoringPresenterFactory<IReorderParametersPresenter>
    {
        private readonly IVBE _vbe;
        private readonly IRefactoringDialog<ReorderParametersViewModel> _view;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public ReorderParametersPresenterFactory(IVBE vbe, IRefactoringDialog<ReorderParametersViewModel> view,
            RubberduckParserState state, IMessageBox messageBox)
        {
            _vbe = vbe;
            _view = view;
            _state = state;
            _messageBox = messageBox;
        }

        public IReorderParametersPresenter Create()
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

            var model = new ReorderParametersModel(_state, selection.Value, _messageBox);
            return new ReorderParametersPresenter(_view, model, _messageBox);

        }
    }
}

using Rubberduck.UI;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldPresenter
    {
        EncapsulateFieldModel Show();
    }

    public class EncapsulateFieldPresenter : IEncapsulateFieldPresenter
    {
        private readonly IEncapsulateFieldView _view;
        private readonly EncapsulateFieldModel _model;
        private readonly IMessageBox _messageBox;

        public EncapsulateFieldPresenter(IEncapsulateFieldView view, EncapsulateFieldModel model, IMessageBox messageBox)
        {
            _view = view;
            _model = model;
            _messageBox = messageBox;
        }

        public EncapsulateFieldModel Show()
        {
            var dialogResult = _view.ShowDialog();

            return null;
        }
    }
}

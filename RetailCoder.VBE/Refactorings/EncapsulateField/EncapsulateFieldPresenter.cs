using System.Windows.Forms;
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
            if (_model.TargetDeclaration == null) { return null; }

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.Name = _view.PropertyName;
            _model.Accessibility = _view.PropertyAccessibility;
            _model.SetterType = _view.PropertySetterType;
            return _model;
        }
    }
}

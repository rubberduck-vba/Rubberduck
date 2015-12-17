using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;

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

        private static readonly string[] PrimitiveTypes =
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.LongPtr,
            Tokens.Integer,
            Tokens.Single,
            Tokens.String,
            Tokens.StrPtr
        };

        public EncapsulateFieldModel Show()
        {
            if (_model.TargetDeclaration == null) { return null; }

            _view.NewPropertyName = _model.TargetDeclaration.IdentifierName;
            _view.TargetDeclaration = _model.TargetDeclaration;
            _view.PropertyAccessibility = EncapsulateFieldDialog.Accessibility.ByVal;
            _view.PropertySetterType = EncapsulateFieldDialog.SetterType.Let;

            if (PrimitiveTypes.Contains(_model.TargetDeclaration.AsTypeName))
            {
                _view.IsPropertySetterTypeChangeable = false;
            }
            else if (_model.TargetDeclaration.AsTypeName != Tokens.Variant)
            {
                _view.PropertySetterType = EncapsulateFieldDialog.SetterType.Set;
                _view.IsPropertySetterTypeChangeable = false;
            }

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.Name = _view.NewPropertyName;
            _model.Accessibility = _view.PropertyAccessibility;
            _model.SetterType = _view.PropertySetterType;
            return _model;
        }
    }
}

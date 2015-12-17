using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;

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

        public EncapsulateFieldPresenter(IEncapsulateFieldView view, EncapsulateFieldModel model)
        {
            _view = view;
            _model = model;
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

            if (PrimitiveTypes.Contains(_model.TargetDeclaration.AsTypeName))
            {
                _view.IsSetterTypeChangeable = false;
            }
            else if (_model.TargetDeclaration.AsTypeName != Tokens.Variant)
            {
                _view.SetterTypeIsLet = false;
                _view.IsSetterTypeChangeable = false;
            }

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.PropertyName = _view.NewPropertyName;
            _model.SetterTypeIsLet = _view.SetterTypeIsLet;

            _model.ParameterName = _view.ParameterName;
            _model.ParameterModifierIsByVal = _view.ParameterModifierIsByVal;
            return _model;
        }
    }
}

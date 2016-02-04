using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
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

            if (_model.TargetDeclaration.References.Any(r => r.IsAssignment))
            {
                if (PrimitiveTypes.Contains(_model.TargetDeclaration.AsTypeName))
                {
                    _view.MustImplementLetSetterType = true;
                    _view.CanImplementSetSetterType = false;
                }
                else if (_model.TargetDeclaration.AsTypeName != Tokens.Variant)
                {
                    _view.MustImplementSetSetterType = true;
                    _view.CanImplementLetSetterType = false;
                }
                else
                {
                    RuleContext node = _model.TargetDeclaration.References.First(r => r.IsAssignment).Context;
                    while (!(node is VBAParser.LetStmtContext) && !(node is VBAParser.SetStmtContext))
                    {
                        node = node.Parent;
                    }

                    if (node is VBAParser.LetStmtContext)
                    {
                        _view.MustImplementLetSetterType = true;
                        _view.CanImplementSetSetterType = false;
                    }
                    else
                    {
                        _view.MustImplementSetSetterType = true;
                        _view.CanImplementLetSetterType = false;
                    }
                }
            }
            else
            {
                if (PrimitiveTypes.Contains(_model.TargetDeclaration.AsTypeName))
                {
                    _view.CanImplementSetSetterType = false;
                }
                else if (_model.TargetDeclaration.AsTypeName != Tokens.Variant)
                {
                    _view.CanImplementLetSetterType = false;
                }
            }

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.PropertyName = _view.NewPropertyName;
            _model.ImplementLetSetterType = _view.MustImplementLetSetterType;
            _model.ImplementSetSetterType = _view.MustImplementSetSetterType;

            _model.ParameterName = _view.ParameterName;
            return _model;
        }
    }
}

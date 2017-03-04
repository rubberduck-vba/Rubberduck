using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldPresenter
    {
        EncapsulateFieldModel Show();
    }

    public class EncapsulateFieldPresenter : IEncapsulateFieldPresenter
    {
        private readonly IEncapsulateFieldDialog _view;
        private readonly EncapsulateFieldModel _model;

        public EncapsulateFieldPresenter(IEncapsulateFieldDialog view, EncapsulateFieldModel model)
        {
            _view = view;
            _model = model;
        }

        public EncapsulateFieldModel Show()
        {
            if (_model.TargetDeclaration == null) { return null; }

            _view.TargetDeclaration = _model.TargetDeclaration;
            _view.NewPropertyName = _model.TargetDeclaration.IdentifierName;

            var isVariant = _model.TargetDeclaration.AsTypeName.Equals(Tokens.Variant);
            var isValueType = !isVariant && (SymbolList.ValueTypes.Contains(_model.TargetDeclaration.AsTypeName) ||
                              _model.TargetDeclaration.DeclarationType == DeclarationType.Enumeration);

            if (_model.TargetDeclaration.References.Any(r => r.IsAssignment))
            {
                if (isVariant)
                {
                    RuleContext node = _model.TargetDeclaration.References.First(r => r.IsAssignment).Context;
                    while (!(node is VBAParser.LetStmtContext) && !(node is VBAParser.SetStmtContext))
                    {
                        node = node.Parent;
                    }

                    if (node is VBAParser.LetStmtContext)
                    {
                        _view.CanImplementLetSetterType = true;
                    }
                    else
                    {
                        _view.CanImplementSetSetterType = true;
                    }                    
                }
                else if (isValueType)
                {
                    _view.CanImplementLetSetterType = true;
                }
                else
                {
                    _view.CanImplementSetSetterType = true;
                }
            }
            else
            {
                if (isValueType)
                {
                    _view.CanImplementLetSetterType = true;
                }
                else if (!isVariant)
                {
                    _view.CanImplementSetSetterType = true;
                }
                else
                {
                    _view.CanImplementLetSetterType = true;
                    _view.CanImplementSetSetterType = true;
                }
            }

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.PropertyName = _view.NewPropertyName;
            _model.ImplementLetSetterType = _view.CanImplementLetSetterType;
            _model.ImplementSetSetterType = _view.CanImplementSetSetterType;
            _model.CanImplementLet = !_view.MustImplementSetSetterType;

            _model.ParameterName = _view.ParameterName;
            return _model;
        }
    }
}

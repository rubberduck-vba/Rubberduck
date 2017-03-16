using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldPresenter
    {
        EncapsulateFieldModel Show();
    }

    public class EncapsulateFieldPresenter : IEncapsulateFieldPresenter
    {
        private readonly IRefactoringDialog<EncapsulateFieldViewModel> _view;
        private readonly EncapsulateFieldModel _model;

        public EncapsulateFieldPresenter(IRefactoringDialog<EncapsulateFieldViewModel> view, EncapsulateFieldModel model)
        {
            _view = view;
            _model = model;
        }

        public EncapsulateFieldModel Show()
        {
            if (_model.TargetDeclaration == null) { return null; }

            _view.ViewModel.TargetDeclaration = _model.TargetDeclaration;

            var isVariant = _model.TargetDeclaration.AsTypeName.Equals(Tokens.Variant);
            var isValueType = !isVariant && (SymbolList.ValueTypes.Contains(_model.TargetDeclaration.AsTypeName) ||
                              _model.TargetDeclaration.DeclarationType == DeclarationType.Enumeration);

            AssignSetterAndLetterAvailability(isVariant, isValueType);

            _view.ShowDialog();
            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }

            _model.PropertyName = _view.ViewModel.PropertyName;
            _model.ImplementLetSetterType = _view.ViewModel.CanHaveLet;
            _model.ImplementSetSetterType = _view.ViewModel.CanHaveSet;
            _model.CanImplementLet = _view.ViewModel.CanHaveSet && !_view.ViewModel.CanHaveSet;

            _model.ParameterName = _view.ViewModel.ParameterName;
            return _model;
        }

        private void AssignSetterAndLetterAvailability(bool isVariant, bool isValueType)
        {
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
                        _view.ViewModel.CanHaveLet = true;
                    }
                    else
                    {
                        _view.ViewModel.CanHaveSet = true;
                    }
                }
                else if (isValueType)
                {
                    _view.ViewModel.CanHaveLet = true;
                }
                else
                {
                    _view.ViewModel.CanHaveSet = true;
                }
            }
            else
            {
                if (isValueType)
                {
                    _view.ViewModel.CanHaveLet = true;
                }
                else if (!isVariant)
                {
                    _view.ViewModel.CanHaveSet = true;
                }
                else
                {
                    _view.ViewModel.CanHaveLet = true;
                    _view.ViewModel.CanHaveSet = true;
                }
            }
        }
    }
}

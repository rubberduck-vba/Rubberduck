using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    internal class EncapsulateFieldPresenter : RefactoringPresenterBase<EncapsulateFieldModel, EncapsulateFieldDialog, EncapsulateFieldView, EncapsulateFieldViewModel>, IEncapsulateFieldPresenter
    {
        public EncapsulateFieldPresenter(EncapsulateFieldModel model,
            IRefactoringDialogFactory dialogFactory, EncapsulateFieldView view) : base(model, dialogFactory, view)
        { }

        public override EncapsulateFieldModel Show()
        {
            if (Model.TargetDeclaration == null) { return null; }

            ViewModel.TargetDeclaration = Model.TargetDeclaration;

            var isVariant = Model.TargetDeclaration.AsTypeName.Equals(Tokens.Variant);
            var isValueType = !isVariant && (SymbolList.ValueTypes.Contains(Model.TargetDeclaration.AsTypeName) ||
                              Model.TargetDeclaration.DeclarationType == DeclarationType.Enumeration);

            AssignSetterAndLetterAvailability(isVariant, isValueType);

            Dialog.ShowDialog();
            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }

            Model.PropertyName = ViewModel.PropertyName;
            Model.ImplementLetSetterType = ViewModel.CanHaveLet;
            Model.ImplementSetSetterType = ViewModel.CanHaveSet;
            Model.CanImplementLet = ViewModel.CanHaveSet && !ViewModel.CanHaveSet;

            Model.ParameterName = ViewModel.ParameterName;
            return Model;
        }

        private void AssignSetterAndLetterAvailability(bool isVariant, bool isValueType)
        {
            if (Model.TargetDeclaration.References.Any(r => r.IsAssignment))
            {
                if (isVariant)
                {
                    RuleContext node = Model.TargetDeclaration.References.First(r => r.IsAssignment).Context;
                    while (!(node is VBAParser.LetStmtContext) && !(node is VBAParser.SetStmtContext))
                    {
                        node = node.Parent;
                    }

                    if (node is VBAParser.LetStmtContext)
                    {
                        ViewModel.CanHaveLet = true;
                    }
                    else
                    {
                        ViewModel.CanHaveSet = true;
                    }
                }
                else if (isValueType)
                {
                    ViewModel.CanHaveLet = true;
                }
                else
                {
                    ViewModel.CanHaveSet = true;
                }
            }
            else
            {
                if (isValueType)
                {
                    ViewModel.CanHaveLet = true;
                }
                else if (!isVariant)
                {
                    ViewModel.CanHaveSet = true;
                }
                else
                {
                    ViewModel.CanHaveLet = true;
                    ViewModel.CanHaveSet = true;
                }
            }
        }
    }
}

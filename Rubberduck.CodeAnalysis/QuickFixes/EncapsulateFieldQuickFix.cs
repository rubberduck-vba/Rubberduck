using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class EncapsulateFieldQuickFix : RefactoringQuickFixBase
    {
        public EncapsulateFieldQuickFix(EncapsulateFieldRefactoring refactoring)
            : base(refactoring, typeof(EncapsulatePublicFieldInspection))
        {}

        protected override void Refactor(IInspectionResult result)
        {
            Refactoring.Refactor(result.Target);
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(Resources.Inspections.QuickFixes.EncapsulatePublicFieldInspectionQuickFix, result.Target.IdentifierName);
        }
    }
}
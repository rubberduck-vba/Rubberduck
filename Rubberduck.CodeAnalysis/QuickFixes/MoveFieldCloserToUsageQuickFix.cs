using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MoveFieldCloserToUsageQuickFix : RefactoringQuickFixBase
    {
        public MoveFieldCloserToUsageQuickFix(MoveCloserToUsageRefactoring refactoring)
            : base(refactoring, typeof(MoveFieldCloserToUsageInspection))
        {}

        protected override void Refactor(IInspectionResult result)
        {
            Refactoring.Refactor(result.Target);
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(InspectionResults.MoveFieldCloserToUsageInspection, result.Target.IdentifierName);
        }
    }
}
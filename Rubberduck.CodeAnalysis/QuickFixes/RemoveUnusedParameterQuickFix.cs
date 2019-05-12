using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Refactorings.RemoveParameters;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnusedParameterQuickFix : RefactoringQuickFixBase
    {
        public RemoveUnusedParameterQuickFix(RemoveParametersRefactoring refactoring)
            : base(refactoring, typeof(ParameterNotUsedInspection))
        {}

        protected override void Refactor(IInspectionResult result)
        {
            ((RemoveParametersRefactoring)Refactoring).QuickFix(result.QualifiedSelection);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnusedParameterQuickFix;
    }
}
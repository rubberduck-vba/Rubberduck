using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete.Refactoring
{
    /// <summary>
    /// Moves field declaration to the procedure scope it's used in.
    /// </summary>
    /// <inspections>
    /// <inspection name="MoveFieldCloserToUsageInspection" />
    /// </inspections>
    /// <canfix multiple="false" procedure="false" module="false" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// Private value As Long
    /// 
    /// Public Sub DoSomething()
    ///     value = 42
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class MoveFieldCloserToUsageQuickFix : RefactoringQuickFixBase
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
            return string.Format(Resources.Inspections.InspectionResults.MoveFieldCloserToUsageInspection, result.Target.IdentifierName);
        }
    }
}
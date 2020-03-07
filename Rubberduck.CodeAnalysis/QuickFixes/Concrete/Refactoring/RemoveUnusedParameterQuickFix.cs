using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Refactorings.RemoveParameters;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Refactors a procedure's signature to remove a parameter that isn't used. Also updates usages.
    /// </summary>
    /// <inspections>
    /// <inspection name="ParameterNotUsedInspection" />
    /// </inspections>
    /// <canfix procedure="false" module="false" project="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     DoSomethingElse 42
    /// End Sub
    /// 
    /// Private Sub DoSomethingElse(ByVal value As Long)
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     DoSomethingElse
    /// End Sub
    /// 
    /// Private Sub DoSomethingElse()
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveUnusedParameterQuickFix : RefactoringQuickFixBase
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
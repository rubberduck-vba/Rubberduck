using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeleteDeclarations;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes the declaration for a variable that is never assigned. This operation may result in broken code if the unassigned variable is in use.
    /// </summary>
    /// <inspections>
    /// <inspection name="VariableNotAssignedInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveUnassignedIdentifierQuickFix : QuickFixBase
    {
        private readonly ICodeOnlyRefactoringAction<DeleteDeclarationsModel> _refactoring;
        public RemoveUnassignedIdentifierQuickFix(DeleteDeclarationsRefactoringAction refactoringAction)
            : base(typeof(VariableNotAssignedInspection))
        {
            _refactoring = refactoringAction;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var model = new DeleteDeclarationsModel(result.Target);

            _refactoring.Refactor(model, rewriteSession);
        }

        public override void Fix(IReadOnlyCollection<IInspectionResult> results, IRewriteSession rewriteSession)
        {
            var model = new DeleteDeclarationsModel(results.Select(r => r.Target));

            _refactoring.Refactor(model, rewriteSession);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnassignedIdentifierQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
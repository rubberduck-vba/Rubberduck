using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adds an explicit Variant data type to an implicitly typed declaration. Note: a more specific data type might be more appropriate.
    /// </summary>
    /// <inspections>
    /// <inspection name="VariableTypeNotDeclaredInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value
    ///     value = Sheet1.Range("A1").Value
    ///     Debug.Print TypeName(value)
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Variant
    ///     value = Sheet1.Range("A1").Value
    ///     Debug.Print TypeName(value)
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class DeclareAsExplicitVariantQuickFix : QuickFixBase
    {
        private readonly ICodeOnlyRefactoringAction<ImplicitTypeToExplicitModel> _refactoring;
        public DeclareAsExplicitVariantQuickFix(ImplicitTypeToExplicitRefactoringAction refactoringAction)
            : base(typeof(VariableTypeNotDeclaredInspection))
        {
            _refactoring = refactoringAction;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var model = new ImplicitTypeToExplicitModel(result.Target)
            {
                ForceVariantAsType = true
            };

            _refactoring.Refactor(model, rewriteSession);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.DeclareAsExplicitVariantQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
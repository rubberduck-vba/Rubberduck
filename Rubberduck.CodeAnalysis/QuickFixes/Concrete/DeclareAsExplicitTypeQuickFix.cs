using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Assigns an explicit data type to an implicitly typed declaration.
    /// </summary>
    /// <inspections>
    /// <inspection name="VariableTypeNotDeclaredInspection" />
    /// <inspection name="ImplicitlyTypedConstInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef arg As String)
    ///     Dim localVar
    ///     localVar = arg
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef arg As String)
    ///     Dim localVar As String
    ///     localVar = arg
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Const PI = 3.14
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Const PI As Double = 3.14
    /// ]]>
    /// </after>
    /// </example>
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething(Optional ByVal arg = 2)
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething(Optional ByVal arg As Long = 2)
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef arg)
    ///     arg = CCur(Sheet1.Range("A1").Value)
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef arg As Currency)
    ///     arg = CCur(Sheet1.Range("A1").Value)
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class DeclareAsExplicitTypeQuickFix : QuickFixBase
    {
        private readonly ICodeOnlyRefactoringAction<ImplicitTypeToExplicitModel> _refactoring;
        public DeclareAsExplicitTypeQuickFix(ImplicitTypeToExplicitRefactoringAction refactoringAction)
            : base(typeof(VariableTypeNotDeclaredInspection), typeof(ImplicitlyTypedConstInspection))
        {
            _refactoring = refactoringAction;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            _refactoring.Refactor(new ImplicitTypeToExplicitModel(result.Target), rewriteSession);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.DeclareAsExplicitTypeQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
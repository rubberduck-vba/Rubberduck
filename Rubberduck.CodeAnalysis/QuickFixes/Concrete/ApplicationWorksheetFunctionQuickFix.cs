using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces a late-bound Application.{Member} call to the corresponding early-bound Application.WorksheetFunction.{Member} call.
    /// </summary>
    /// <inspections>
    /// <inspection name="ApplicationWorksheetFunctionInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Debug.Print Application.Sum(Sheet1.Range("A1:A10"))
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Debug.Print Application.WorksheetFunction.Sum(Sheet1.Range("A1:A10"))
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ApplicationWorksheetFunctionQuickFix : QuickFixBase
    {
        public ApplicationWorksheetFunctionQuickFix()
            : base(typeof(ApplicationWorksheetFunctionInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "WorksheetFunction.");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ApplicationWorksheetFunctionQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}

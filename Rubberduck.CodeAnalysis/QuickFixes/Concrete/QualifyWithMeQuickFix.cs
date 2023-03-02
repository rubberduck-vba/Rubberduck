using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Qualifies an implicit reference with 'Me'.
    /// </summary>
    /// <inspections>
    /// <inspection name="ImplicitContainingWorksheetReferenceInspection" />
    /// <inspection name="ImplicitContainingWorkbookReferenceInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Range("A1")
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Me.Range("A1")
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal class QualifyWithMeQuickFix : QuickFixBase
    {
        public QualifyWithMeQuickFix()
            : base(typeof(ImplicitContainingWorkbookReferenceInspection),
                typeof(ImplicitContainingWorksheetReferenceInspection))
        {}
        
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var context = result.Context;
            rewriter.InsertBefore(context.Start.TokenIndex, $"{Tokens.Me}.");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.QualifyWithMeQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
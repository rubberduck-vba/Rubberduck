using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces 'While...Wend' loop statement with equivalent 'Do While...Loop'.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObsoleteWhileWendStatementInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     While value < 99
    ///         value = value + 1
    ///     Wend
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     Do While value < 99
    ///         value = value + 1
    ///     Loop
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ReplaceWhileWendWithDoWhileLoopQuickFix : QuickFixBase
    {
        public ReplaceWhileWendWithDoWhileLoopQuickFix()
            : base(typeof(ObsoleteWhileWendStatementInspection))
        { }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.WhileWendStmtContext)result.Context;

            rewriter.Replace(context.WHILE(), "Do While");
            rewriter.Replace(context.WEND(), "Loop");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ReplaceWhileWendWithDoWhileLoopQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}

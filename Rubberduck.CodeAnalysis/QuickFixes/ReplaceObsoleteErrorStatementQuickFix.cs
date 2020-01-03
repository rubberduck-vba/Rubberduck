using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Replaces the obsolete 'Error' statement with an 'ErrObject.Raise' member call.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObsoleteErrorSyntaxInspection" />
    /// </inspections>
    /// <canfix procedure="true" module="true" project="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Error 5
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Err.Raise 5
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    public sealed class ReplaceObsoleteErrorStatementQuickFix : QuickFixBase
    {
        public ReplaceObsoleteErrorStatementQuickFix()
            : base(typeof(ObsoleteErrorSyntaxInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.ErrorStmtContext) result.Context;

            rewriter.Replace(context.ERROR(), "Err.Raise");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ReplaceObsoleteErrorStatementQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
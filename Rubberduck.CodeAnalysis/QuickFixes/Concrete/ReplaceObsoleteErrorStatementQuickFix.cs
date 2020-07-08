using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces the obsolete 'Error' statement with an 'ErrObject.Raise' member call.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObsoleteErrorSyntaxInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
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
    internal sealed class ReplaceObsoleteErrorStatementQuickFix : QuickFixBase
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

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
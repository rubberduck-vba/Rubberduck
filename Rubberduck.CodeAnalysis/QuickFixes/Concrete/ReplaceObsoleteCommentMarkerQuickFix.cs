using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces the obsolete 'Rem' comment marker token with a single quote character.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObsoleteCommentSyntaxInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Rem some comment...
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     ' some comment...
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ReplaceObsoleteCommentMarkerQuickFix : QuickFixBase
    {
        public ReplaceObsoleteCommentMarkerQuickFix()
            : base(typeof(ObsoleteCommentSyntaxInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.RemCommentContext) result.Context;

            rewriter.Replace(context.REM(), "'");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveObsoleteStatementQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Removes a comment.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObsoleteCommentSyntaxInspection" />
    /// </inspections>
    /// <canfix procedure="true" module="true" project="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Rem does something
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveCommentQuickFix : QuickFixBase
    {
        public RemoveCommentQuickFix()
            : base(typeof(ObsoleteCommentSyntaxInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Remove(result.Context);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveCommentQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
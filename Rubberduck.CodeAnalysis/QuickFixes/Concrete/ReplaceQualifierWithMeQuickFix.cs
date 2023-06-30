using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Resources;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces an explicit qualifier with 'Me'.
    /// </summary>
    /// <inspections>
    /// <inspection name="SuspiciousPredeclaredInstanceAccessInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// Private ClickCount As Long
    /// 
    /// Private Sub CommandButton1_Click()
    ///     ClickCount = ClickCount + 1
    ///     ' works fine as long as the current instance is the default instance
    ///     UserForm1.TextBox1.Text = ClickCount
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// Private ClickCount As Long
    /// 
    /// Private Sub CommandButton1_Click()
    ///     ClickCount = ClickCount + 1
    ///     ' works fine regardless of which instance we're in
    ///     Me.TextBox1.Text = ClickCount
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal class ReplaceQualifierWithMeQuickFix : QuickFixBase
    {
        public ReplaceQualifierWithMeQuickFix()
            :base(typeof(SuspiciousPredeclaredInstanceAccessInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var context = result.Context;
            rewriter.Replace(context.Start, Tokens.Me);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ReplaceQualifierWithMeQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}

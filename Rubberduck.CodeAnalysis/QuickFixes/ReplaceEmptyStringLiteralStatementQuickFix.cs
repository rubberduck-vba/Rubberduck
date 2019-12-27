using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Replaces empty string literals '""' with the 'vbNullString' constant.
    /// </summary>
    /// <inspections>
    /// <inspection name="EmptyStringLiteralInspection" />
    /// </inspections>
    /// <canfix procedure="true" module="true" project="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     If ActiveCell.Text = "" Then
    ///         Debug.Print ActiveCell.Address
    ///     End If
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     If ActiveCell.Text = vbNullString Then
    ///         Debug.Print ActiveCell.Address
    ///     End If
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    public sealed class ReplaceEmptyStringLiteralStatementQuickFix : QuickFixBase
    {
        public ReplaceEmptyStringLiteralStatementQuickFix()
            : base(typeof(EmptyStringLiteralInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Replace(result.Context, "vbNullString");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.EmptyStringLiteralInspectionQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
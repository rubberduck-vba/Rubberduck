using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces empty string literals '""' with the 'vbNullString' constant.
    /// </summary>
    /// <inspections>
    /// <inspection name="EmptyStringLiteralInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
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
    internal sealed class ReplaceEmptyStringLiteralStatementQuickFix : QuickFixBase
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

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
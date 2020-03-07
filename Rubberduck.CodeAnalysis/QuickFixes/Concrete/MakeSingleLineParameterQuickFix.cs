using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Rewrites a parameter declaration that is split across multiple lines.
    /// </summary>
    /// <inspections>
    /// <inspection name="MultilineParameterInspection" />
    /// </inspections>
    /// <canfix procedure="true" module="true" project="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal value _
    ///     As Long)
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal value As Long)
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class MakeSingleLineParameterQuickFix : QuickFixBase
    {
        public MakeSingleLineParameterQuickFix()
            : base(typeof(MultilineParameterInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var parameter = result.Context.GetText()
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            rewriter.Replace(result.Context, parameter);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.MakeSingleLineParameterQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}

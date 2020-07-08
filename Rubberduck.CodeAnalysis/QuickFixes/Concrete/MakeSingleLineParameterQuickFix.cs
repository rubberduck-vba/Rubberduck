using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Rewrites a parameter declaration that is split across multiple lines.
    /// </summary>
    /// <inspections>
    /// <inspection name="MultilineParameterInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
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

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}

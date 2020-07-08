using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces 'Global' access modifier with the equivalent 'Public' keyword.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObsoleteGlobalInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// Global Something As Long
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// Public Something As Long
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ReplaceGlobalModifierQuickFix : QuickFixBase
    {
        public ReplaceGlobalModifierQuickFix()
            : base(typeof(ObsoleteGlobalInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.Replace(((ParserRuleContext)result.Context.Parent.Parent).GetDescendent<VBAParser.VisibilityContext>(), Tokens.Public);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ObsoleteGlobalInspectionQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Replaces the Dim keyword with Private in a module-scoped declaration.
    /// </summary>
    /// <inspections>
    /// <inspection name="ModuleScopeDimKeywordInspection" />
    /// </inspections>
    /// <canfix procedure="false" module="true" project="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Dim thing As Long
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print thing
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Private thing As Long
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print thing
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    public sealed class ChangeDimToPrivateQuickFix : QuickFixBase
    {
        public ChangeDimToPrivateQuickFix()
            : base(typeof(ModuleScopeDimKeywordInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var context = (VBAParser.VariableStmtContext)result.Context.Parent.Parent;
            rewriter.Replace(context.DIM(), Tokens.Private);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ChangeDimToPrivateQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
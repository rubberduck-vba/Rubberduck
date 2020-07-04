using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces the Dim keyword with Private in a module-scoped declaration.
    /// </summary>
    /// <inspections>
    /// <inspection name="ModuleScopeDimKeywordInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="true" project="true" all="true" />
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
    internal sealed class ChangeDimToPrivateQuickFix : QuickFixBase
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

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
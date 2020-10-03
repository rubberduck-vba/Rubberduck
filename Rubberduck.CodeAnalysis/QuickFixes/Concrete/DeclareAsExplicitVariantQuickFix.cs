using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adds an explicit Variant data type to an implicitly typed declaration. Note: a more specific data type might be more appropriate.
    /// </summary>
    /// <inspections>
    /// <inspection name="VariableTypeNotDeclaredInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value
    ///     value = Sheet1.Range("A1").Value
    ///     Debug.Print TypeName(value)
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Variant
    ///     value = Sheet1.Range("A1").Value
    ///     Debug.Print TypeName(value)
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class DeclareAsExplicitVariantQuickFix : QuickFixBase
    {
        public DeclareAsExplicitVariantQuickFix()
            : base(typeof(VariableTypeNotDeclaredInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);

            ParserRuleContext identifierNode =
                result.Context is VBAParser.VariableSubStmtContext || result.Context is VBAParser.ConstSubStmtContext
                ? result.Context.children[0]
                : ((dynamic) result.Context).unrestrictedIdentifier();
            rewriter.InsertAfter(identifierNode.Stop.TokenIndex, " As Variant");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.DeclareAsExplicitVariantQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
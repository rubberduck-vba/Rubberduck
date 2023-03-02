using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Replaces the use of a Variant-returning standard library function with its String-returning equivalent. Using this when the argument can be 'Null' will break the code.
    /// </summary>
    /// <inspections>
    /// <inspection name="UntypedFunctionUsageInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print Format(VBA.DateTime.Date, "yyyy-MM-dd")
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print Format$(VBA.DateTime.Date, "yyyy-MM-dd")
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class UntypedFunctionUsageQuickFix : QuickFixBase
    {
        public UntypedFunctionUsageQuickFix()
            : base(typeof(UntypedFunctionUsageInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertAfter(result.Context.Stop.TokenIndex, "$");
        }

        private static string GetNewSignature(ParserRuleContext context)
        {
            Debug.Assert(context != null);

            return context.children.Aggregate(string.Empty, (current, member) =>
            {
                var isIdentifierNode = member is VBAParser.IdentifierContext;
                return current + member.GetText() + (isIdentifierNode ? "$" : string.Empty);
            });
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(Resources.Inspections.QuickFixes.UseTypedFunctionQuickFix, result.Context.GetText(), GetNewSignature(result.Context));
        }

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
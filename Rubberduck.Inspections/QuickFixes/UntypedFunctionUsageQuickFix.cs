using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class UntypedFunctionUsageQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public UntypedFunctionUsageQuickFix(RubberduckParserState state)
            : base(typeof(UntypedFunctionUsageInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertAfter(result.Context.Stop.TokenIndex, "$");
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(Resources.Inspections.QuickFixes.UseTypedFunctionQuickFix, result.Context.GetText(), GetNewSignature(result.Context));
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

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
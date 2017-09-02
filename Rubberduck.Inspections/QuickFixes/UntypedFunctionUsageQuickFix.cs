using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class UntypedFunctionUsageQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public UntypedFunctionUsageQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<UntypedFunctionUsageInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertAfter(result.Context.Stop.TokenIndex, "$");
        }

        public string Description(IInspectionResult result)
        {
            return string.Format(InspectionsUI.QuickFixUseTypedFunction_, result.Context.GetText(), GetNewSignature(result.Context));
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

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}
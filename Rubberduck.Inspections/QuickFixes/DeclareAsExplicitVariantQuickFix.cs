using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class DeclareAsExplicitVariantQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(VariableTypeNotDeclaredInspection)
        };
        
        public DeclareAsExplicitVariantQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);

            ParserRuleContext identifierNode = result.Context is VBAParser.VariableSubStmtContext
                ? result.Context.children[0]
                : ((dynamic) result.Context).unrestrictedIdentifier();
            rewriter.InsertAfter(identifierNode.Stop.TokenIndex, " As Variant");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.DeclareAsExplicitVariantQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}
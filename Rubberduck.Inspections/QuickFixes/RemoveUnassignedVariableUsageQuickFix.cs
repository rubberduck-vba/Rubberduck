using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnassignedVariableUsageQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(UnassignedVariableUsageInspection)
        };

        public RemoveUnassignedVariableUsageQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var assignmentContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(result.Context) ??
                                                  (ParserRuleContext)ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(result.Context);

            rewriter.Remove(assignmentContext);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveUnassignedVariableUsageQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}
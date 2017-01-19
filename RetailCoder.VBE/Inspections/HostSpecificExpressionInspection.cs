using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class HostSpecificExpressionInspection : InspectionBase
    {
        public HostSpecificExpressionInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.HostSpecificExpressionInspectionMeta; } }
        public override string Description { get { return InspectionsUI.HostSpecificExpressionInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            return Declarations.Where(item => item.DeclarationType == DeclarationType.BracketedExpression)
                .Select(item => new HostSpecificExpressionInspectionResult(this, item)).ToList();
        }
    }
}
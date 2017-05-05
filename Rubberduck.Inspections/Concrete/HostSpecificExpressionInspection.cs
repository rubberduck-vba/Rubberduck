using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class HostSpecificExpressionInspection : InspectionBase
    {
        public HostSpecificExpressionInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Declarations.Where(item => item.DeclarationType == DeclarationType.BracketedExpression)
                .Select(item => new DeclarationInspectionResult(this, string.Format(InspectionsUI.HostSpecificExpressionInspectionResultFormat, item.IdentifierName), item));
        }
    }
}
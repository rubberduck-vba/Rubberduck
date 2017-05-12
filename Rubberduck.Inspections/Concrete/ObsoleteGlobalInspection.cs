using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObsoleteGlobalInspection : InspectionBase
    {
        public ObsoleteGlobalInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where item.Accessibility == Accessibility.Global && item.Context != null
                         select new DeclarationInspectionResult(this,
                                                     string.Format(InspectionsUI.ObsoleteGlobalInspectionResultFormat, item.DeclarationType.ToLocalizedString(), item.IdentifierName),
                                                     item);

            return issues;
        }
    }
}

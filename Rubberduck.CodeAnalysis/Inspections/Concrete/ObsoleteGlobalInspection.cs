using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObsoleteGlobalInspection : InspectionBase
    {
        public ObsoleteGlobalInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where item.Accessibility == Accessibility.Global && item.Context != null
                         select new DeclarationInspectionResult(this,
                                                     string.Format(InspectionResults.ObsoleteGlobalInspection, item.DeclarationType.ToLocalizedString(), item.IdentifierName),
                                                     item);

            return issues;
        }
    }
}

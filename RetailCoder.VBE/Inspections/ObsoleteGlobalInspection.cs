using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteGlobalInspection : InspectionBase
    {
        public ObsoleteGlobalInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteGlobalInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteGlobalInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where item.Accessibility == Accessibility.Global && item.Context != null
                         select new ObsoleteGlobalInspectionResult(this, new QualifiedContext<ParserRuleContext>(item.QualifiedName.QualifiedModuleName, item.Context));

            return issues;
        }
    }
}

using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteGlobalInspection : InspectionBase
    {
        public ObsoleteGlobalInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public override string Description { get { return RubberduckUI.ObsoleteGlobal; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where item.Accessibility == Accessibility.Global && item.Context != null
                         select new ObsoleteGlobalInspectionResult(this, Description, new QualifiedContext<ParserRuleContext>(item.QualifiedName.QualifiedModuleName, item.Context));

            return issues;
        }
    }
}
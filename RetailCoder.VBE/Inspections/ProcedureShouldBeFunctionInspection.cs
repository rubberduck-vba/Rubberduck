using System.Collections.Generic;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public class ProcedureShouldBeFunctionInspection : IInspection
    {
        public ProcedureShouldBeFunctionInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ProcedureShouldBeFunctionInspection"; } }
        public string Description { get { return InspectionsUI.ProcedureShouldBeFunctionInspection; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            return new List<CodeInspectionResultBase>();
        }
    }
}

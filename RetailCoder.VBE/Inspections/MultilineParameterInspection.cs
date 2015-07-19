using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class MultilineParameterInspection : IInspection 
    {
        public MultilineParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "MultilineParameterInspection"; } }
        public string Description { get { return RubberduckUI.MultilineParameter_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var multilineParameters = from p in parseResult.Declarations.Items
                .Where(item => item.DeclarationType == DeclarationType.Parameter)
                where p.Context.GetSelection().LineCount > 1
                select p;

            var issues = multilineParameters
                .Select(param => new MultilineParameterInspectionResult(string.Format(param.Context.GetSelection().LineCount > 3 ? RubberduckUI.EasterEgg_Continuator : Description, param.IdentifierName), Severity, param.Context, param.QualifiedName));

            return issues;
        }
    }
}
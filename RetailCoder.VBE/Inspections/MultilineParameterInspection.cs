using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return RubberduckUI.MultilineParameter_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var multilineParameters = from p in state.AllUserDeclarations
                .Where(item => item.DeclarationType == DeclarationType.Parameter)
                where p.Context.GetSelection().LineCount > 1
                select p;

            var issues = multilineParameters
                .Select(param => new MultilineParameterInspectionResult(this, string.Format(param.Context.GetSelection().LineCount > 3 ? RubberduckUI.EasterEgg_Continuator : Description, param.IdentifierName), param.Context, param.QualifiedName));

            return issues;
        }
    }
}
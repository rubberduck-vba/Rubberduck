using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspection : IInspection
    {
        public OptionBaseInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.OptionBase; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            foreach (var module in parseResult)
            {
                var declarationLines = module.Component.CodeModule.CountOfDeclarationLines;
                if (declarationLines == 0)
                {
                    declarationLines = 1;
                }

                var lines = module.Component.CodeModule.get_Lines(1, declarationLines).Split('\n')
                    .Select(line => line.Replace("\r", string.Empty));
                var option = Tokens.Option + " " + Tokens.Base + " 1";
                if (lines.Contains(option))
                {
                    yield return new OptionBaseInspectionResult(Name, Severity, module.QualifiedName);
                }
            }
        }
    }
}
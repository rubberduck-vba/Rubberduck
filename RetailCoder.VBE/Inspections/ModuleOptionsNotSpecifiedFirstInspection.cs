using System.Collections.Generic;
using System.Linq;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class ModuleOptionsNotSpecifiedFirstInspection : IInspection
    {
        public ModuleOptionsNotSpecifiedFirstInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ModuleOptionsNotSpecifiedFirst; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            foreach (var module in parseResult.ComponentParseResults)
            {
                var declarationLines = module.Component.CodeModule.CountOfDeclarationLines;
                if (declarationLines == 0)
                {
                    declarationLines = 1;
                }

                var lines = GetIndexedOptionLines(module.Component.CodeModule.get_Lines(1, declarationLines).Split('\n')
                    .Select(line => line.Replace("\r", string.Empty)).ToArray());

                if (lines.Any() && lines.Count != lines.Keys.Max() + 1)
                {
                    // todo: figure this one out
                    yield return new ModuleOptionsNotSpecifiedFirstInspectionResult(Name, Severity, new CommentNode(string.Empty, new QualifiedSelection(module.QualifiedName, Selection.Empty)));
                }
            }
        }

        private IDictionary<int, string> GetIndexedOptionLines(string[] lines)
        {
            var result = new Dictionary<int, string>();
            for (var i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith(Tokens.Option))
                {
                    result.Add(i, lines[i]);
                }
            }

            return result;
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class ModuleOptionsNotSpecifiedFirstInspection //: IInspection // disabled
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

                var commentLines = module.Comments
                    .Where(node => node.QualifiedSelection.Selection.StartColumn == 1
                        && node.QualifiedSelection.Selection.StartLine <= declarationLines)
                    .Select(node => node.QualifiedSelection.Selection);

                var lines = GetIndexedOptionLines(module.Component.CodeModule.get_Lines(1, declarationLines).Split('\n')
                    .Select(line => line.Replace("\r", string.Empty)).Where(line => !string.IsNullOrEmpty(line.Trim())).ToArray());

                if (lines.Any() && lines.Count != lines.Keys.Max() + 1)
                {
                    // todo: figure this one out
                    yield return new ModuleOptionsNotSpecifiedFirstInspectionResult(Name, Severity, new CommentNode(string.Empty, new QualifiedSelection(module.QualifiedName, Selection.Home)));
                }
            }
        }

        private IDictionary<int, string> GetIndexedOptionLines(string[] lines)
        {
            var result = new Dictionary<int, string>();
            for (var i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith(Tokens.Option) || string.IsNullOrEmpty(lines[i]))
                {
                    var trimmed = lines[i].TrimStart();
                    if (trimmed.StartsWith(Tokens.CommentMarker) || trimmed.StartsWith(Tokens.Rem))
                    {
                        result.Add(i, lines[i]);
                    }
                }
            }

            return result;
        }
    }
}
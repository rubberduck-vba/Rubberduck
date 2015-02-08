using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class OptionExplicitInspection : IInspection
    {
        public OptionExplicitInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.OptionExplicit; } }
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
                                                       .Select(line => line.Replace("\r",string.Empty));
                var option = Tokens.Option + " " + Tokens.Explicit;
                if (!lines.Contains(option))
                {
                    yield return new OptionExplicitInspectionResult(Name, Severity, module.QualifiedName);
                }
            }
        }
    }
}
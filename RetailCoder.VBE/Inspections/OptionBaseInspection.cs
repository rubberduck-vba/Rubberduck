using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspection : IInspection
    {
        public OptionBaseInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.OptionBase; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var result = new List<CodeInspectionResultBase>();
            foreach (var module in parseResult.ComponentParseResults)
            {
                try
                {
                    var declarationLines = module.Component.CodeModule.CountOfDeclarationLines;
                    if (declarationLines == 0)
                    {
                        declarationLines = 1;
                    }

                    if (module.Component.CodeModule.CountOfLines > 0)
                    {
                        var lines = module.Component.CodeModule.get_Lines(1, declarationLines).Split('\n')
                            .Select(line => line.Replace("\r", string.Empty));
                        var option = Tokens.Option + " " + Tokens.Base + " 1";
                        if (lines.Contains(option))
                        {
                            result.Add(new OptionBaseInspectionResult(Name, Severity, module.QualifiedName));
                        }
                    }
                }
                catch (COMException)
                {
                    // couldn't access the CodeModule. Whiskey Tango Foxtrot.
                }
            }

            return result;
        }
    }
}
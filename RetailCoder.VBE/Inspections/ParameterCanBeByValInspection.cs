using System.Collections.Generic;
using Rubberduck.Parsing;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspection : IInspection
    {
        public ParameterCanBeByValInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ParameterCanBeByVal_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = 
                parseResult.Declarations.Items.Where(declaration =>
                    declaration.DeclarationType == DeclarationType.Parameter
                    && ((VBAParser.ArgContext) declaration.Context).BYVAL() == null
                    && !declaration.References.Any(reference => reference.IsAssignment));

            foreach (var issue in declarations)
            {
                yield return new ParameterCanBeByValInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName);
            }
        }
    }
}
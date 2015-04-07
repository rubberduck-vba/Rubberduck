using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class AssignedByValParameterInspection //: IInspection /* note: deferred to v1.4 */
    {
        public AssignedByValParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ByValParameterIsAssigned_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var assignedByValParameters =
                parseResult.Declarations.Items.Where(declaration =>
                    declaration.DeclarationType == DeclarationType.Parameter
                    && ((VBAParser.ArgContext)declaration.Context).BYVAL() != null
                    && declaration.References.Any(reference => reference.IsAssignment));

            var issues = assignedByValParameters
                .Select(param => new AssignedByValParameterInspectionResult(string.Format(Name, param.IdentifierName), Severity, param.Context, param.QualifiedName));

            return issues;
        }
    }
}
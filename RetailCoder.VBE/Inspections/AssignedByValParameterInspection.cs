using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class AssignedByValParameterInspection : IInspection
    {
        public AssignedByValParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "AssignedByValParameterInspection"; } }
        public string Description { get { return RubberduckUI.ByValParameterIsAssigned_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private string AnnotationName { get { return Name.Replace("Inspection", string.Empty); } }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState parseResult)
        {
            var name = AnnotationName;
            var assignedByValParameters =
                parseResult.AllDeclarations.Where(declaration => !declaration.IsInspectionDisabled(name)
                    && !declaration.IsBuiltIn 
                    && declaration.DeclarationType == DeclarationType.Parameter
                    && ((VBAParser.ArgContext)declaration.Context).BYVAL() != null
                    && declaration.References.Any(reference => reference.IsAssignment));

            var issues = assignedByValParameters
                .Select(param => new AssignedByValParameterInspectionResult(this, string.Format(Description, param.IdentifierName), param.Context, param.QualifiedName));

            return issues;
        }
    }
}
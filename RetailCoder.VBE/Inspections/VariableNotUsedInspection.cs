using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class VariableNotUsedInspection : IInspection
    {
        public VariableNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "VariableNotUsedInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return RubberduckUI.VariableNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var declarations = state.AllDeclarations.Where(declaration =>
                !declaration.IsBuiltIn 
                && !declaration.IsWithEvents
                && declaration.DeclarationType == DeclarationType.Variable
                && declaration.References.All(reference => reference.IsAssignment));

            return declarations.Select(issue => 
                new IdentifierNotUsedInspectionResult(this, issue, ((dynamic)issue.Context).ambiguousIdentifier(), issue.QualifiedName.QualifiedModuleName));
        }
    }
}
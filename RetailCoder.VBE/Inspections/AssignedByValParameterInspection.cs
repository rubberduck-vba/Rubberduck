using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class AssignedByValParameterInspection : InspectionBase
    {
        public AssignedByValParameterInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.ByValParameterIsAssigned_; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var assignedByValParameters = UserDeclarations.Where(declaration => 
                    declaration.DeclarationType == DeclarationType.Parameter
                    && ((VBAParser.ArgContext)declaration.Context).BYVAL() != null
                    && declaration.References.Any(reference => reference.IsAssignment));

            var issues = assignedByValParameters
                .Select(param => new AssignedByValParameterInspectionResult(this, string.Format(Description, param.IdentifierName), param.Context, param.QualifiedName));

            return issues;
        }
    }
}
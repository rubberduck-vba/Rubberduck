using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class AssignedByValParameterInspection : InspectionBase
    {
        public AssignedByValParameterInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = DefaultSeverity;
        }

        public override string Meta { get { return InspectionsUI.AssignedByValParameterInspectionMeta; } }
        public override string Description { get { return InspectionsUI.AssignedByValParameterInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var parameters = State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .Where(item => !IsIgnoringInspectionResultFor(item, AnnotationName))
                .Where(item => ((VBAParser.ArgContext) item.Context).BYVAL() != null
                               && item.References.Any(reference => reference.IsAssignment));

            return parameters
                .Select(param => new AssignedByValParameterInspectionResult(this, param))
                .ToList();
        }
    }
}

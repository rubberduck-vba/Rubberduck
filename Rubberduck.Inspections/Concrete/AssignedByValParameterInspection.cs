using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class AssignedByValParameterInspection : InspectionBase
    {
        public AssignedByValParameterInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = DefaultSeverity;
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var parameters = State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .Cast<ParameterDeclaration>()
                .Where(item => !item.IsByRef 
                    && !IsIgnoringInspectionResultFor(item, AnnotationName)
                    && item.References.Any(reference => reference.IsAssignment));

            return parameters
                .Select(param => new DeclarationInspectionResult(this,
                                                      string.Format(InspectionsUI.AssignedByValParameterInspectionResultFormat, param.IdentifierName),
                                                      param));
        }
    }
}

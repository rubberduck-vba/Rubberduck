using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.Inspections
{
    public sealed class AssignedByValParameterInspection : InspectionBase
    {
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private readonly RubberduckParserState _parserState;
        public AssignedByValParameterInspection(RubberduckParserState state, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(state)
        {
            Severity = DefaultSeverity;
            _dialogFactory = dialogFactory;
            _parserState = state;

        }

        public override string Meta { get { return InspectionsUI.AssignedByValParameterInspectionMeta; } }
        public override string Description { get { return InspectionsUI.AssignedByValParameterInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var parameters = State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .OfType<ParameterDeclaration>()
                .Where(item => !item.IsByRef 
                    && !IsIgnoringInspectionResultFor(item, AnnotationName)
                    && item.References.Any(reference => reference.IsAssignment))
                .ToList();

            return parameters
                .Select(param => new AssignedByValParameterInspectionResult(this, param, _parserState, _dialogFactory))
                .ToList();
        }
    }
}

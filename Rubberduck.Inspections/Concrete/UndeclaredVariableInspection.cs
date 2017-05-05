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
    public sealed class UndeclaredVariableInspection : InspectionBase
    {
        public UndeclaredVariableInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => item.IsUndeclared && !IsIgnoringInspectionResultFor(item, AnnotationName))
                .Select(item => new DeclarationInspectionResult(this, string.Format(InspectionsUI.UndeclaredVariableInspectionResultFormat, item.IdentifierName), item));
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UndeclaredVariableInspection : InspectionBase
    {
        public UndeclaredVariableInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => item.IsUndeclared && !IsIgnoringInspectionResultFor(item, AnnotationName))
                .Select(item => new DeclarationInspectionResult(this, string.Format(InspectionResults.UndeclaredVariableInspection, item.IdentifierName), item));
        }
    }
}
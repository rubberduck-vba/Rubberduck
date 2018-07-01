using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UndeclaredRedimVariableInspection : InspectionBase
    {
        public UndeclaredRedimVariableInspection(RubberduckParserState state) 
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(item => item.IsRedimVariable && !IsIgnoringInspectionResultFor(item, AnnotationName))
                .Select(item => new DeclarationInspectionResult(this, string.Format(InspectionResults.UndeclaredRedimVariableInspection, item.IdentifierName), item));
        }
    }
}

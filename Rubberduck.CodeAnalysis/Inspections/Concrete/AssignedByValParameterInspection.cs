using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class AssignedByValParameterInspection : InspectionBase
    {
        public AssignedByValParameterInspection(RubberduckParserState state)
            : base(state)
        { }
        
        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var parameters = State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .Cast<ParameterDeclaration>()
                .Where(item => !item.IsByRef 
                    && !item.IsIgnoringInspectionResultFor(AnnotationName)
                    && item.References.Any(reference => reference.IsAssignment));

            return parameters
                .Select(param => new DeclarationInspectionResult(this,
                                                      string.Format(InspectionResults.AssignedByValParameterInspection, param.IdentifierName),
                                                      param));
        }
    }
}

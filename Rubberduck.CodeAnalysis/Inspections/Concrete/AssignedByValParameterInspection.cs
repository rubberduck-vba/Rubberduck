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
    /// <summary>
    /// Warns about parameters passed by value being assigned a new value in the body of a procedure.
    /// </summary>
    /// <why>
    /// 
    /// </why>
    /// <example>
    /// This inspection means to flag the following examples:
    /// <code>
    /// </code>
    /// The following code should not trip this inspection:
    /// <code>
    /// </code>
    /// </example>
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

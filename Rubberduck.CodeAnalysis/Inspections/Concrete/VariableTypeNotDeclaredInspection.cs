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
    public sealed class VariableTypeNotDeclaredInspection : InspectionBase
    {
        public VariableTypeNotDeclaredInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var issues = from item in State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                         .Union(State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter))
                         where (item.DeclarationType != DeclarationType.Parameter || (item.DeclarationType == DeclarationType.Parameter && !item.IsArray))
                         && item.DeclarationType != DeclarationType.Control
                         && !item.IsIgnoringInspectionResultFor(AnnotationName)
                         && !item.IsTypeSpecified
                         && !item.IsUndeclared
                         select new DeclarationInspectionResult(this, string.Format(InspectionResults.ImplicitVariantDeclarationInspection, item.DeclarationType, item.IdentifierName), item);
            return issues;
        }
    }
}

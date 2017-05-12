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
    public sealed class VariableTypeNotDeclaredInspection : InspectionBase
    {
        public VariableTypeNotDeclaredInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var issues = from item in State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                         .Union(State.DeclarationFinder.UserDeclarations(DeclarationType.Constant))
                         .Union(State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter))
                         where (item.DeclarationType != DeclarationType.Parameter || (item.DeclarationType == DeclarationType.Parameter && !item.IsArray))
                         && item.DeclarationType != DeclarationType.Control
                         && !IsIgnoringInspectionResultFor(item, AnnotationName)
                         && !item.IsTypeSpecified
                         && !item.IsUndeclared
                         select new DeclarationInspectionResult(this,
                                                     string.Format(InspectionsUI.ImplicitVariantDeclarationInspectionResultFormat,
                                                                   item.DeclarationType,
                                                                   item.IdentifierName),
                                                     item);

            return issues;
        }
    }
}

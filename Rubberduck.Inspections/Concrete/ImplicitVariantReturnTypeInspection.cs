using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitVariantReturnTypeInspection : InspectionBase
    {
        public ImplicitVariantReturnTypeInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var issues = from item in State.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                         where !item.IsTypeSpecified && !IsIgnoringInspectionResultFor(item, AnnotationName)
                         let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                         select new DeclarationInspectionResult(this,
                                                     string.Format(InspectionsUI.ImplicitVariantReturnTypeInspectionResultFormat, item.IdentifierName),
                                                     item);
            return issues;
        }
    }
}

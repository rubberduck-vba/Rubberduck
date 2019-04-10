using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitVariantReturnTypeInspection : InspectionBase
    {
        public ImplicitVariantReturnTypeInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var issues = from item in State.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                         where !item.IsTypeSpecified && !item.IsIgnoringInspectionResultFor(AnnotationName)
                         let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                         select new DeclarationInspectionResult(this,
                                                     string.Format(InspectionResults.ImplicitVariantReturnTypeInspection, item.IdentifierName),
                                                     item);
            return issues;
        }
    }
}

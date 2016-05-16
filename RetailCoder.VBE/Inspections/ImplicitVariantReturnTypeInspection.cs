using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitVariantReturnTypeInspection : InspectionBase
    {
        public ImplicitVariantReturnTypeInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.ImplicitVariantReturnTypeInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitVariantReturnTypeInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.LibraryFunction
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                               where !item.IsInspectionDisabled(AnnotationName)
                                && ProcedureTypes.Contains(item.DeclarationType)
                                && !item.IsTypeSpecified()
                               let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                               select new ImplicitVariantReturnTypeInspectionResult(this, issue.Declaration.IdentifierName, issue.QualifiedContext);
            return issues;
        }
    }
}

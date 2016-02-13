using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitVariantReturnTypeInspection : InspectionBase
    {
        public ImplicitVariantReturnTypeInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.ImplicitVariantReturnType_; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.LibraryFunction
        };

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                               where !item.IsInspectionDisabled(AnnotationName)
                                && ProcedureTypes.Contains(item.DeclarationType)
                                && !item.IsTypeSpecified()
                               let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                               select new ImplicitVariantReturnTypeInspectionResult(this, string.Format(Description, issue.Declaration.IdentifierName), issue.QualifiedContext);
            return issues;
        }
    }
}
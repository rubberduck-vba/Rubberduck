using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspection : IInspection
    {
        public ImplicitVariantReturnTypeInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ImplicitVariantReturnTypeInspection"; } }
        public string Description { get { return RubberduckUI.ImplicitVariantReturnType_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.LibraryFunction
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = from item in parseResult.Declarations.Items
                               where !item.IsBuiltIn
                                && ProcedureTypes.Contains(item.DeclarationType)
                                && !item.IsTypeSpecified()
                               let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                               select new ImplicitVariantReturnTypeInspectionResult(string.Format(Description, issue.Declaration.IdentifierName), Severity, issue.QualifiedContext);
            return issues;
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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

        private string AnnotationName { get { return Name.Replace("Inspection", string.Empty); } }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState parseResult)
        {
            var issues = from item in parseResult.AllDeclarations
                               where !item.IsInspectionDisabled(AnnotationName)
                                && !item.IsBuiltIn
                                && ProcedureTypes.Contains(item.DeclarationType)
                                && !item.IsTypeSpecified()
                               let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                               select new ImplicitVariantReturnTypeInspectionResult(this, string.Format(Description, issue.Declaration.IdentifierName), issue.QualifiedContext);
            return issues;
        }
    }
}
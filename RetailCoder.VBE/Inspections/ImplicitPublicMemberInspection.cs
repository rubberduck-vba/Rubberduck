using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ImplicitPublicMemberInspection : IInspection
    {
        public ImplicitPublicMemberInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ImplicitPublicMemberInspection"; } }
        public string Description { get { return RubberduckUI.ImplicitPublicMember_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private string AnnotationName { get { return Name.Replace("Inspection", string.Empty); } }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState parseResult)
        {
            var issues = from item in parseResult.AllDeclarations
                         where !item.IsInspectionDisabled(AnnotationName) 
                               && !item.IsBuiltIn
                               && ProcedureTypes.Contains(item.DeclarationType)
                               && item.Accessibility == Accessibility.Implicit
                         let context = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)
                               select new ImplicitPublicMemberInspectionResult(this, string.Format(Description, ((dynamic)context.Context).ambiguousIdentifier().GetText()), context);
            return issues;
        }
    }
}
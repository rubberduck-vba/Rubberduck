using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ImplicitPublicMemberInspection : IInspection
    {
        public ImplicitPublicMemberInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ImplicitPublicMember_; } }
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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = from item in parseResult.Declarations.Items
                         where ProcedureTypes.Contains(item.DeclarationType)
                               && item.Accessibility == Accessibility.Implicit
                         let context = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)
                               select new ImplicitPublicMemberInspectionResult(string.Format(Name, ((dynamic)context.Context).ambiguousIdentifier().GetText()), Severity, context);
            return issues;
        }
    }
}
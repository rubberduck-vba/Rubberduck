using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Listeners;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Listeners;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ImplicitPublicMemberInspection : IInspection
    {
        public ImplicitPublicMemberInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
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
            var declarations = from item in parseResult.Declarations.Items
                               where ProcedureTypes.Contains(item.DeclarationType)
                               && item.Accessibility == Accessibility.Implicit
                               select new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context.Parent as ParserRuleContext);

            foreach (var declaration in declarations)
            {
                yield return new ImplicitPublicMemberInspectionResult(string.Format(Name, ((dynamic)declaration.Context).ambiguousIdentifier().GetText()), Severity, declaration);
            }
        }
    }
}
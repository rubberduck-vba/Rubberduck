using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitPublicMemberInspection : InspectionBase
    {
        public ImplicitPublicMemberInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.ImplicitPublicMember_; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where !item.IsInspectionDisabled(AnnotationName) 
                               && ProcedureTypes.Contains(item.DeclarationType)
                               && item.Accessibility == Accessibility.Implicit
                         let context = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)
                               select new ImplicitPublicMemberInspectionResult(this, string.Format(Description, ((dynamic)context.Context).ambiguousIdentifier().GetText()), context);
            return issues;
        }
    }
}
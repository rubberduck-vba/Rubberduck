using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitPublicMemberInspection : InspectionBase
    {
        public ImplicitPublicMemberInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where ProcedureTypes.Contains(item.DeclarationType)
                               && item.Accessibility == Accessibility.Implicit
                         let context = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)
                         select new InspectionResult(this,
                                                     string.Format(InspectionsUI.ImplicitPublicMemberInspectionResultFormat, item.IdentifierName),
                                                     context,
                                                     item);
            return issues;
        }
    }
}

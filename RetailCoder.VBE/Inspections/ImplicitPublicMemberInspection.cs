using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitPublicMemberInspection : InspectionBase
    {
        public ImplicitPublicMemberInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.ImplicitPublicMemberInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitPublicMemberInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where ProcedureTypes.Contains(item.DeclarationType)
                               && item.Accessibility == Accessibility.Implicit
                         let context = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)
                         select new ImplicitPublicMemberInspectionResult(this, context, item);
            return issues;
        }
    }
}

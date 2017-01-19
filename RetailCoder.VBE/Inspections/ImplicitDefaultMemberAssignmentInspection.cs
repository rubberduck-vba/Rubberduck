using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ImplicitDefaultMemberAssignmentInspection : InspectionBase
    {
        public ImplicitDefaultMemberAssignmentInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var interestingDeclarations =
                State.AllDeclarations.Where(item =>
                    item.AsTypeDeclaration != null
                    && ClassModuleDeclaration.HasDefaultMember(item.AsTypeDeclaration));

            var interestingReferences = interestingDeclarations
                .SelectMany(declaration => declaration.References)
                .Where(reference =>
                {
                    var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
                    return reference.IsAssignment && letStmtContext != null && letStmtContext.LET() == null;
                });

            return interestingReferences.Select(reference => new ImplicitDefaultMemberAssignmentInspectionResult(this, reference));
        }

        public override string Meta { get { return InspectionsUI.ImplicitDefaultMemberAssignmentInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ImplicitDefaultMemberAssignmentInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
    }
}
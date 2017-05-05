using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitDefaultMemberAssignmentInspection : InspectionBase
    {
        public ImplicitDefaultMemberAssignmentInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
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

            return interestingReferences.Select(reference => new IdentifierReferenceInspectionResult(this,
                                                                                  string.Format(InspectionsUI.ImplicitDefaultMemberAssignmentInspectionResultFormat,
                                                                                                reference.Declaration.IdentifierName,
                                                                                                reference.Declaration.AsTypeDeclaration.IdentifierName),
                                                                                  State,
                                                                                  reference));
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;
    }
}
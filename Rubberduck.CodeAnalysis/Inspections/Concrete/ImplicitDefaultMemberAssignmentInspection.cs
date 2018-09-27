using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitDefaultMemberAssignmentInspection : InspectionBase
    {
        public ImplicitDefaultMemberAssignmentInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var interestingDeclarations =
                State.AllDeclarations.Where(item =>
                    item.AsTypeDeclaration != null
                    && ClassModuleDeclaration.HasDefaultMember(item.AsTypeDeclaration));

            var interestingReferences = interestingDeclarations
                .SelectMany(declaration => declaration.References)
                .Where(reference =>
                {
                    var letStmtContext = reference.Context.GetAncestor<VBAParser.LetStmtContext>();
                    return reference.IsAssignment 
                           && letStmtContext != null 
                           && letStmtContext.LET() == null
                           && !reference.IsIgnoringInspectionResultFor(AnnotationName);
                });

            return interestingReferences.Select(reference => new IdentifierReferenceInspectionResult(this,
                                                                                  string.Format(InspectionResults.ImplicitDefaultMemberAssignmentInspection,
                                                                                                reference.Declaration.IdentifierName,
                                                                                                reference.Declaration.AsTypeDeclaration.IdentifierName),
                                                                                  State,
                                                                                  reference));
        }
    }
}
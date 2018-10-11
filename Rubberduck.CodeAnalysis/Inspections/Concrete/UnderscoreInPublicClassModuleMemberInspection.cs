using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UnderscoreInPublicClassModuleMemberInspection : InspectionBase
    {
        public UnderscoreInPublicClassModuleMemberInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var names = State.DeclarationFinder.UserDeclarations(Parsing.Symbols.DeclarationType.Member)
                .Where(w => w.ParentDeclaration.DeclarationType == Parsing.Symbols.DeclarationType.ClassModule)
                .Where(w => w.Accessibility == Parsing.Symbols.Accessibility.Public)
                .Where(w => w.IdentifierName.Contains('_'))
                .ToList();

            return names.Select(issue =>
                new DeclarationInspectionResult(this, string.Format(InspectionResults.UnderscoreInPublicClassModuleMemberInspection, issue.IdentifierName), issue));
        }
    }
}

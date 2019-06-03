using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// This inspection means to indicate when the project has not been renamed.
    /// </summary>
    /// <why>
    /// VBA projects should be meaningfully named, to avoid namespace clashes when referencing other VBA projects.
    /// </why>
    [CannotAnnotate]
    public sealed class DefaultProjectNameInspection : InspectionBase
    {
        public DefaultProjectNameInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var projects = State.DeclarationFinder.UserDeclarations(DeclarationType.Project)
                .Where(item => item.IdentifierName.StartsWith("VBAProject"))
                .ToList();

            return projects
                .Select(issue => new DeclarationInspectionResult(this, Description, issue))
                .ToList();
        }
    }
}

using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// This inspection means to indicate when the project has not been renamed.
    /// </summary>
    /// <why>
    /// VBA projects should be meaningfully named, to avoid namespace clashes when referencing other VBA projects.
    /// </why>
    [CannotAnnotate]
    public sealed class DefaultProjectNameInspection : DeclarationInspectionBase
    {
        public DefaultProjectNameInspection(RubberduckParserState state)
            : base(state, DeclarationType.Project)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;

            var results = new List<IInspectionResult>();
            foreach (var projectDeclaration in State.DeclarationFinder.UserDeclarations(DeclarationType.Project))
            {
                if (projectDeclaration == null)
                {
                    continue;
                }

                var module = projectDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module, finder));
            }

            return results;
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.IdentifierName.StartsWith("VBAProject");
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return Description;
        }
    }
}

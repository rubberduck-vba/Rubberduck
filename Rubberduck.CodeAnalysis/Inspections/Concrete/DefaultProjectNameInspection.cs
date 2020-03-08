using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// This inspection means to indicate when the project has not been renamed.
    /// </summary>
    /// <why>
    /// VBA projects should be meaningfully named, to avoid namespace clashes when referencing other VBA projects.
    /// </why>
    [CannotAnnotate]
    internal sealed class DefaultProjectNameInspection : DeclarationInspectionBase
    {
        public DefaultProjectNameInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Project)
        {}

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

using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class ImplicitTypeInspectionBase : DeclarationInspectionBase
    {
        protected ImplicitTypeInspectionBase(IDeclarationFinderProvider declarationFinderProvider, params DeclarationType[] relevantDeclarationTypes) 
            : base(declarationFinderProvider, relevantDeclarationTypes)
        {}

        protected ImplicitTypeInspectionBase(IDeclarationFinderProvider declarationFinderProvider, DeclarationType[] relevantDeclarationTypes, DeclarationType[] excludeDeclarationTypes)
            : base(declarationFinderProvider, relevantDeclarationTypes, excludeDeclarationTypes)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return !declaration.IsTypeSpecified;
        }
    }
}
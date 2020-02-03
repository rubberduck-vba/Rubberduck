using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class ImplicitTypeInspectionBase : DeclarationInspectionBase
    {
        protected ImplicitTypeInspectionBase(RubberduckParserState state, params DeclarationType[] relevantDeclarationTypes) 
        :base(state, relevantDeclarationTypes)
        { }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return !declaration.IsTypeSpecified;
        }
    }
}
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public interface IDeclarationFinderProvider
    {
        DeclarationFinder DeclarationFinder { get; }

        void RefreshDeclarationFinder();
    }
}

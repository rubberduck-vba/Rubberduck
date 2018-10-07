using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.VBA
{
    public interface IDeclarationFinderProvider
    {
        DeclarationFinder DeclarationFinder { get; }

        void RefreshDeclarationFinder();
    }
}

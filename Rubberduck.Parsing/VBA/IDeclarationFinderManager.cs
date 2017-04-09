

namespace Rubberduck.Parsing.VBA
{
    public interface IDeclarationFinderManager : IDeclarationFinderProvider 
    {
        void RefreshDeclarationFinder();
    }
}

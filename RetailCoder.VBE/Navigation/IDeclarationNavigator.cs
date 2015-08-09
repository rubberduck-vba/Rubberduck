using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Navigation
{
    public interface IDeclarationNavigator
    {
        void Find();
        void Find(Declaration target);
    }
}
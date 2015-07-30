using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Navigations
{
    public interface IFindAllReferences
    {
        void Find();
        void Find(Declaration target);
    }
}
using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA
{
    public interface ICommonTokenStreamProvider
    {
        CommonTokenStream Tokens(string code); 
    }
}

using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public interface ICommonTokenStreamProvider
    {
        CommonTokenStream Tokens(string code); 
    }
}

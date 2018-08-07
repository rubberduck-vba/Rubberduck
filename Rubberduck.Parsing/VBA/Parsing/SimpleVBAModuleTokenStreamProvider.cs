using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class SimpleVBAModuleTokenStreamProvider : ICommonTokenStreamProvider
    {
        public CommonTokenStream Tokens(string code)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            return new CommonTokenStream(lexer);
        }
    }
}

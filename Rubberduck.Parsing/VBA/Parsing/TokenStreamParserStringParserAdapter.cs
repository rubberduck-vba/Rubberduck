using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class TokenStreamParserStringParserAdapter : IStringParser
    {
        private readonly ICommonTokenStreamProvider _tokenStreamProvider;
        private readonly ITokenStreamParser _tokenStreamParser;

        public TokenStreamParserStringParserAdapter(ICommonTokenStreamProvider tokenStreamProvider, ITokenStreamParser tokenStreamParser)
        {
            _tokenStreamProvider = tokenStreamProvider;
            _tokenStreamParser = tokenStreamParser;
        }

        public (IParseTree tree, ITokenStream tokenStream) Parse(string moduleName, string projectId, string code, CancellationToken token,
            CodeKind codeKind = CodeKind.SnippetCode, ParserMode parserMode = ParserMode.FallBackSllToLl)
        {
            token.ThrowIfCancellationRequested();
            var tokenStream = _tokenStreamProvider.Tokens(code);
            token.ThrowIfCancellationRequested();
            var tree = _tokenStreamParser.Parse(moduleName, tokenStream, codeKind, parserMode);
            return (tree, tokenStream);
        }
    }
}

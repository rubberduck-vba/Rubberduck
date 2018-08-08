using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.PreProcessing;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class TokenStreamParserStringParserAdapterWithPreprocessing : IStringParser
    {
        private readonly ICommonTokenStreamProvider _tokenStreamProvider;
        private readonly ITokenStreamParser _tokenStreamParser;
        private readonly ITokenStreamPreprocessor _preprocessor;

        public TokenStreamParserStringParserAdapterWithPreprocessing(ICommonTokenStreamProvider tokenStreamProvider, ITokenStreamParser tokenStreamParser, ITokenStreamPreprocessor preprocessor)
        {
            _tokenStreamProvider = tokenStreamProvider;
            _tokenStreamParser = tokenStreamParser;
            _preprocessor = preprocessor;
        }

        public (IParseTree tree, ITokenStream tokenStream) Parse(string moduleName, string projectId, string code, CancellationToken token,
            CodeKind codeKind = CodeKind.SnippetCode, ParserMode parserMode = ParserMode.FallBackSllToLl)
        {
            token.ThrowIfCancellationRequested();
            var tokenStream = _tokenStreamProvider.Tokens(code);
            token.ThrowIfCancellationRequested();
            tokenStream = _preprocessor.PreprocessTokenStream(projectId, moduleName, tokenStream, token, codeKind);
            token.ThrowIfCancellationRequested();
            var tree = _tokenStreamParser.Parse(moduleName, tokenStream, codeKind, parserMode);
            return (tree, tokenStream);
        }
    }
}

using Antlr4.Runtime;
using System.Threading;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPreprocessor : ITokenStreamPreprocessor
    {
        private readonly double _vbaVersion;
        private readonly ITokenStreamParser _parser;
        private readonly ICompilationArgumentsProvider _compilationArgumentsProvider;

        public VBAPreprocessor(double vbaVersion, ITokenStreamParser preprocessorParser, ICompilationArgumentsProvider compilationArgumentsProvider)
        {
            _vbaVersion = vbaVersion;
            _compilationArgumentsProvider = compilationArgumentsProvider;
            _parser = preprocessorParser;
        }

        public CommonTokenStream PreprocessTokenStream(string projectId, string moduleName, CommonTokenStream tokenStream, CancellationToken token, CodeKind codeKind = CodeKind.SnippetCode)
        {
            token.ThrowIfCancellationRequested();

            var tree = _parser.Parse(moduleName, tokenStream, codeKind);
            token.ThrowIfCancellationRequested();

            var charStream = tokenStream.TokenSource.InputStream;
            var symbolTable = new SymbolTable<string, IValue>();
            var userCompilationArguments = _compilationArgumentsProvider.UserDefinedCompilationArguments(projectId);
            var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion), userCompilationArguments, charStream, tokenStream);
            var expr = evaluator.Visit(tree);
            var processedTokens = expr.Evaluate(); //This does the actual preprocessing of the token stream as a side effect.
            tokenStream.Reset();
            return tokenStream;
        }
    }
}

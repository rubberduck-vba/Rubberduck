using Antlr4.Runtime;
using System.Threading;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPreprocessor : ITokenStreamPreprocessor
    {
        private readonly ITokenStreamParser _parser;
        private readonly ICompilationArgumentsProvider _compilationArgumentsProvider;

        public VBAPreprocessor(ITokenStreamParser preprocessorParser, ICompilationArgumentsProvider compilationArgumentsProvider)
        {
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
            var predefinedCompilationArgument = _compilationArgumentsProvider.PredefinedCompilationConstants; 
            var evaluator = new VBAPreprocessorVisitor(symbolTable, predefinedCompilationArgument, userCompilationArguments, charStream, tokenStream);
            var expr = evaluator.Visit(tree);
            var processedTokens = expr.Evaluate(); //This does the actual preprocessing of the token stream as a side effect.
            tokenStream.Reset();
            return tokenStream;
        }
    }
}

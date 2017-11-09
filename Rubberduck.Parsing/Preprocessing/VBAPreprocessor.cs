using Antlr4.Runtime;
using System.Threading;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPreprocessor : IVBAPreprocessor
    {
        private readonly double _vbaVersion;
        private readonly VBAPrecompilationParser _parser;

        public VBAPreprocessor(double vbaVersion)
        {
            _vbaVersion = vbaVersion;
            _parser = new VBAPrecompilationParser();
        }

        public void PreprocessTokenStream(string moduleName, CommonTokenStream tokenStream, BaseErrorListener errorListener, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var symbolTable = new SymbolTable<string, IValue>();
            var tree = _parser.Parse(moduleName, tokenStream, errorListener);
            token.ThrowIfCancellationRequested();
            var stream = tokenStream.TokenSource.InputStream;
            var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion), stream, tokenStream);
            var expr = evaluator.Visit(tree);
            var processedTokens = expr.Evaluate(); //This does the actual preprocessing of the token stream as a side effect.
            tokenStream.Reset();
        }
    }
}

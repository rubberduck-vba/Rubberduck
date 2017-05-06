using Antlr4.Runtime;
using System.Threading;

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

        public CommonTokenStream Execute(string moduleName, CommonTokenStream unprocessedTokenStream, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var symbolTable = new SymbolTable<string, IValue>();
            var tree = _parser.Parse(moduleName, unprocessedTokenStream);
            token.ThrowIfCancellationRequested();
            var stream = unprocessedTokenStream.TokenSource.InputStream;
            var evaluator = new VBATokenStreamPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion), stream, unprocessedTokenStream);
            var expr = evaluator.Visit(tree);
            var dummyValue = expr.Evaluate();
            unprocessedTokenStream.Reset();
            return unprocessedTokenStream;
        }
    }
}

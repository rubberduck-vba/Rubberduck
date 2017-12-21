using System.Linq;
using Antlr4.Runtime;
using System.Threading;
using Antlr4.Runtime.Atn;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPreprocessor : IVBAPreprocessor
    {
        private readonly RubberduckParserState _state;
        private readonly double _vbaVersion;
        private readonly VBAPrecompilationParser _parser;

        public VBAPreprocessor(RubberduckParserState state, double vbaVersion)
        {
            _state = state;
            _vbaVersion = vbaVersion;
            _parser = new VBAPrecompilationParser();
        }

        // we need both because sometimes we use the component name and sometimes the fully-qualified name
        public void PreprocessTokenStream(QualifiedModuleName module, string moduleName, CommonTokenStream tokenStream, BaseErrorListener errorListener, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var symbolTable = new SymbolTable<string, IValue>();

            var tree = _parser.Parse(moduleName, PredictionMode.Sll, tokenStream, errorListener);
            if (_state.ModuleExceptions.Any(r => r.Item1 == module))
            {
                tree = _parser.Parse(moduleName, PredictionMode.Ll, tokenStream, errorListener);
            }

            token.ThrowIfCancellationRequested();
            var stream = tokenStream.TokenSource.InputStream;
            var evaluator = new VBAPreprocessorVisitor(_state, module, symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion), stream, tokenStream);
            var expr = evaluator.Visit(tree);
            var processedTokens = expr.Evaluate(); //This does the actual preprocessing of the token stream as a side effect.
            tokenStream.Reset();
        }
    }
}

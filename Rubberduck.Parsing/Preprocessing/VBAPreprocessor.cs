namespace Rubberduck.Parsing.Preprocessing
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

        public string Execute(string moduleName, string unprocessedCode)
        {
            return Preprocess(moduleName, unprocessedCode);
        }

        private string Preprocess(string moduleName, string unprocessedCode)
        {
            SymbolTable<string, IValue> symbolTable = new SymbolTable<string, IValue>();
            var tree = _parser.Parse(moduleName, unprocessedCode);
            var stream = tree.Start.InputStream;
            var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion), stream);
            var expr = evaluator.Visit(tree);
            return expr.Evaluate().AsString;
        }
    }
}

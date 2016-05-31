using System;

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
            try
            {
                return Preprocess(moduleName, unprocessedCode);
            }
            catch (Exception ex)
            {
                throw new VBAPreprocessorException("Exception encountered during preprocessing.", ex);
            }
        }

        private string Preprocess(string moduleName, string unprocessedCode)
        {
            SymbolTable<string, IValue> symbolTable = new SymbolTable<string, IValue>();
            var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion));
            var tree = _parser.Parse(moduleName, unprocessedCode);
            var expr = evaluator.Visit(tree);
            return expr.Evaluate().AsString;
        }
    }
}

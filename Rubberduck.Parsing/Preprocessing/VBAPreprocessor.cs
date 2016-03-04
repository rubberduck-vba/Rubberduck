using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class VBAPreprocessor
    {
        private readonly double _vbaVersion;

        public VBAPreprocessor(double vbaVersion)
        {
            _vbaVersion = vbaVersion;
        }

        public string Execute(string unprocessedCode)
        {
            try
            {
                return Preprocess(unprocessedCode);
            }
            catch (Exception ex)
            {
                throw new VBAPreprocessorException("Exception encountered during preprocessing.", ex);
            }
        }

        private string Preprocess(string unprocessedCode)
        {
            SymbolTable<string, IValue> symbolTable = new SymbolTable<string, IValue>();
            var stream = new AntlrInputStream(unprocessedCode);
            var lexer = new VBAConditionalCompilationLexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAConditionalCompilationParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion));
            var tree = parser.compilationUnit();
            var expr = evaluator.Visit(tree);
            return expr.Evaluate().AsString;
        }
    }
}

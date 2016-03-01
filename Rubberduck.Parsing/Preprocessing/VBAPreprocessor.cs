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

        public string Process(string unprocessedCode)
        {
            try
            {
                SymbolTable symbolTable = new SymbolTable();
                var stream = new AntlrInputStream(unprocessedCode);
                var lexer = new VBAConditionalCompilationLexer(stream);
                var tokens = new CommonTokenStream(lexer);
                var parser = new VBAConditionalCompilationParser(tokens);
                parser.AddErrorListener(new ExceptionErrorListener());
                var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion));
                var tree = parser.compilationUnit();
                var preprocessedModuleBody = (string)evaluator.Visit(tree);
                return preprocessedModuleBody;
            }
            catch (Exception ex)
            {
                throw new VBAPreprocessorException("Exception encountered during preprocessing.", ex);
            }
        }
    }
}

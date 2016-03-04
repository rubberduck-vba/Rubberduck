using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
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
                var optionCompare = ExtractOptionCompare(unprocessedCode);
                return Preprocess(unprocessedCode, optionCompare);
            }
            catch (Exception ex)
            {
                throw new VBAPreprocessorException("Exception encountered during preprocessing.", ex);
            }
        }

        private VBAOptionCompare ExtractOptionCompare(string unprocessedCode)
        {
            var stream = new AntlrInputStream(unprocessedCode);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            var optionCompareListener = new OptionCompareListener();
            parser.AddParseListener(optionCompareListener);
            var tree = parser.startRule();
            return optionCompareListener.OptionCompare;
        }

        private string Preprocess(string unprocessedCode, VBAOptionCompare optionCompare)
        {
            SymbolTable<string, IValue> symbolTable = new SymbolTable<string, IValue>();
            var stream = new AntlrInputStream(unprocessedCode);
            var lexer = new VBAConditionalCompilationLexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAConditionalCompilationParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            var evaluator = new VBAPreprocessorVisitor(symbolTable, new VBAPredefinedCompilationConstants(_vbaVersion), optionCompare);
            var tree = parser.compilationUnit();
            var expr = evaluator.Visit(tree);
            return expr.Evaluate().AsString;
        }
    }
}

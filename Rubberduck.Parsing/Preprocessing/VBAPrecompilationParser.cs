using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class VBAPrecompilationParser
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public VBAConditionalCompilationParser.CompilationUnitContext Parse(string moduleName, string unprocessedCode)
        {
            var stream = new AntlrInputStream(unprocessedCode);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAConditionalCompilationParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            VBAConditionalCompilationParser.CompilationUnitContext tree = null;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.compilationUnit();
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "SLL mode failed in module {0}. Retrying using LL.", moduleName);
                tokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.compilationUnit();
            }
            return tree;
        }
    }
}

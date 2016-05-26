using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBAModuleParser
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public IParseTree Parse(string moduleName, string moduleCode, IParseTreeListener[] listeners, out ITokenStream outStream)
        {
            var stream = new AntlrInputStream(moduleCode);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            foreach (var listener in listeners)
            {
                parser.AddParseListener(listener);
            }
            ParserRuleContext tree = null;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.startRule();
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "SLL mode failed in module {0}. Retrying using LL.", moduleName);
                tokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.startRule();
            }
            outStream = tokens;
            return tree;
        }
    }
}

using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Diagnostics;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBAModuleParser
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public IParseTree Parse(string module, IParseTreeListener[] listeners, out ITokenStream outStream)
        {
            var stream = new AntlrInputStream(module);
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
            catch
            {
                _logger.Warn("{0}: SLL mode failed. Retrying using LL.", GetType().Name);
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

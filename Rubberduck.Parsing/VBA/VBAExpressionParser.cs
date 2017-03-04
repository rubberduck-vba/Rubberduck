using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using NLog;
using Rubberduck.Parsing.Grammar;
using System;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBAExpressionParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Parses the given VBA expression.
        /// </summary>
        /// <param name="expression">The expression to parse. NOTE: Call statements are not supported.</param>
        /// <returns>The root of the parse tree.</returns>
        public ParserRuleContext Parse(string expression)
        {
            var stream = new AntlrInputStream(expression);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            //parser.AddErrorListener(new ExceptionErrorListener());
            ParserRuleContext tree;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.expression();
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "SLL mode failed for {0}. Retrying using LL.", expression);
                tokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.expression();
            }
            return tree;
        }
    }
}

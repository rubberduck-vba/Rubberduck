using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Grammar;
using System;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBAModuleParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public IParseTree Parse(string moduleName, string moduleCode, IParseTreeListener[] listeners, BaseErrorListener errorListener, out ITokenStream outStream)
        {
            var stream = new AntlrInputStream(moduleCode);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAParser(tokens);
            parser.AddErrorListener(errorListener);
            ParserRuleContext tree;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.startRule();
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "SLL mode failed in module {0}. Retrying using LL.", moduleName);
                tokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.startRule();
            }
            foreach (var listener in listeners)
            {
                ParseTreeWalker.Default.Walk(listener, tree);
            }
            outStream = tokens;
            return tree;
        }
    }
}

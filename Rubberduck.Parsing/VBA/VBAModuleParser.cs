using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBAModuleParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public (IParseTree tree, ITokenStream tokenStream) Parse(string moduleName, PredictionMode predictionMode, CommonTokenStream moduleTokens, IParseTreeListener[] listeners, BaseErrorListener errorListener)
        {
            moduleTokens.Reset();
            var parser = new VBAParser(moduleTokens);
            parser.AddErrorListener(errorListener);
            parser.ErrorHandler = new RecoveryStrategy();
            parser.Interpreter.PredictionMode = predictionMode;
            ParserRuleContext tree = parser.startRule();
            foreach (var listener in listeners)
            {
                ParseTreeWalker.Default.Walk(listener, tree);
            }
            return (tree, moduleTokens);
        }
    }
}

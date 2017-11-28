using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Grammar;
using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.ParsingExceptions;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBAModuleParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public (IParseTree tree, ITokenStream tokenStream) Parse(string moduleName, CommonTokenStream moduleTokens, IParseTreeListener[] listeners, BaseErrorListener errorListener)
        {
            moduleTokens.Reset();
            var parser = new VBAParser(moduleTokens);
            parser.AddErrorListener(errorListener);
            ParserRuleContext tree;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.startRule();
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                var parsePassText = syntaxErrorException.ParsePass == ParsePass.CodePanePass
                    ? "code pane"
                    : "exported";
                Logger.Warn($"SLL mode failed while parsing the {parsePassText} version of module {moduleName} at symbol {syntaxErrorException.OffendingSymbol.Text} at L{syntaxErrorException.LineNumber}C{syntaxErrorException.Position}. Retrying using LL.");
                Logger.Debug(syntaxErrorException, "SLL mode exception");
                moduleTokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.startRule();
            }
            catch (Exception exception)
            {
                Logger.Warn($"SLL mode failed while parsing module {moduleName}. Retrying using LL.");
                Logger.Debug(exception, "SLL mode exception");
                moduleTokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.startRule();
            }
            foreach (var listener in listeners)
            {
                ParseTreeWalker.Default.Walk(listener, tree);
            }
            return (tree, moduleTokens);
        }
    }
}

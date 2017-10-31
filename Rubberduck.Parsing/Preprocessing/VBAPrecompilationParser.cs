using System;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPrecompilationParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public VBAConditionalCompilationParser.CompilationUnitContext Parse(string moduleName, CommonTokenStream unprocessedTokenStream, BaseErrorListener errorListener)
        {
            unprocessedTokenStream.Reset();
            var parser = new VBAConditionalCompilationParser(unprocessedTokenStream);
            parser.AddErrorListener(errorListener);
            VBAConditionalCompilationParser.CompilationUnitContext tree;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.compilationUnit();
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                var parsePassText = syntaxErrorException.ParsePass == ParsePass.CodePanePass
                    ? "code pane"
                    : "exported";
                Logger.Warn($"SLL mode failed while preprocessing the {parsePassText} version of module {moduleName} at symbol {syntaxErrorException.OffendingSymbol.Text} at L{syntaxErrorException.LineNumber}C{syntaxErrorException.Position}. Retrying using LL.");
                Logger.Debug(syntaxErrorException, "SLL mode exception");
                unprocessedTokenStream.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.compilationUnit();
            }
            catch (Exception exception)
            {
                Logger.Warn($"SLL mode failed while prprocessing module {moduleName}. Retrying using LL.");
                Logger.Debug(exception, "SLL mode exception");
                unprocessedTokenStream.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.compilationUnit();
            }
            return tree;
        }
    }
}

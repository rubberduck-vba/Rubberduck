using System;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Grammar
{
    /// <summary>
    /// The class is meant to be used for where we need to do rewriting without actually
    /// rewriting an existing module in a VBA project. The result is completely free-floating
    /// and it is the caller's responsibility to persist the finalized results into one of
    /// actual code. This is useful for creating a code preview for refactoring.
    /// </summary>
    public class VBACodeStringParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly CommonTokenStream tokenStream;
        public TokenStreamRewriter Rewriter { get => new TokenStreamRewriter(tokenStream); }
        public IParseTree ParseTree { get; }

        /// <summary>
        /// Parse a given string representing a module code. The code passed in must be
        /// a valid module body, containing valid declarations or complete procedures. 
        /// Code snippets are not valid. 
        /// </summary>
        /// <param name="moduleName">For logging purpose, provide descritpive name for where the virtual module is used</param>
        /// <param name="moduleCodeString">A string containing valid module body.</param>
        public VBACodeStringParser(string moduleName, string moduleCodeString)
        {
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            tokenStream = tokenStreamProvider.Tokens(moduleCodeString);
            var parser = new VBAParser(tokenStream);
            parser.AddErrorListener(new MainParseExceptionErrorListener(moduleName, ParsePass.CodePanePass));
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                ParseTree = parser.startRule();
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                var parsePassText = syntaxErrorException.ParsePass == ParsePass.CodePanePass
                    ? "code pane"
                    : "exported";
                Logger.Warn($"SLL mode failed while parsing the {parsePassText} version of module {moduleName} at symbol {syntaxErrorException.OffendingSymbol.Text} at L{syntaxErrorException.LineNumber}C{syntaxErrorException.Position}. Retrying using LL.");
                Logger.Debug(syntaxErrorException, "SLL mode exception");
                tokenStream.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                ParseTree = parser.startRule();
            }
            catch (Exception exception)
            {
                Logger.Warn($"SLL mode failed while parsing module {moduleName}. Retrying using LL.");
                Logger.Debug(exception, "SLL mode exception");
                tokenStream.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                ParseTree = parser.startRule();
            }
        }
    }
}

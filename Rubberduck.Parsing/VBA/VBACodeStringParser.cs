using System;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols.ParsingExceptions;

namespace Rubberduck.Parsing.VBA
{
    /// <summary>
    /// The class is meant to be used for where we need to do rewriting without actually
    /// rewriting an existing module in a VBA project. The result is completely free-floating
    /// and it is the caller's responsibility to persist the finalized results into one of
    /// actual code. This is useful for creating a code preview for refactoring.
    /// </summary>
    public class VBACodeStringParser
    {
        public enum ParserMode
        {
            Default,
            Sll,
            Ll
        }

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly CommonTokenStream _tokenStream;
        private readonly VBAParser _parser;
        private readonly string _moduleName;
        private readonly ParserMode _mode;

        /// <summary>
        /// Parse a given string representing a module code. The code passed in must be
        /// a valid module body, containing valid declarations or complete procedures. 
        /// Code snippets are not valid. 
        /// </summary>
        /// <param name="moduleName">For logging purpose, provide descritpive name for where the virtual module is used</param>
        /// <param name="moduleCodeString">A string containing valid module body.</param>
        /// <param name="mode">Indicates what parser mode to use. By default we use SLL fallbacking to LL. When mode is explicitly specified, only that mode will be used with no fallback.</param>
        public VBACodeStringParser(string moduleName, string moduleCodeString, ParserMode mode = ParserMode.Default)
        {
            var tokenStreamProvider = new SimpleVBAModuleTokenStreamProvider();
            _tokenStream = tokenStreamProvider.Tokens(moduleCodeString);
            _parser = new VBAParser(_tokenStream);
            _parser.AddErrorListener(new MainParseExceptionErrorListener(moduleName, ParsePass.CodePanePass));
            _moduleName = moduleName;
            _mode = mode;
        }

        public (IParseTree parseTree, TokenStreamRewriter rewriter) Parse()
        {
            var parseTree = _mode == ParserMode.Default ? ParseInternal(_moduleName) : ParseInternal(_mode);
            var rewriter = new TokenStreamRewriter(_tokenStream);

            return (parseTree, rewriter);
        }

        private IParseTree ParseInternal(string moduleName)
        {
            try
            {
                return ParseInternal(ParserMode.Sll);
            }
            catch (ParsePassSyntaxErrorException syntaxErrorException)
            {
                var parsePassText = syntaxErrorException.ParsePass == ParsePass.CodePanePass
                    ? "code pane"
                    : "exported";
                var message = $"SLL mode failed while parsing the {parsePassText} version of module {moduleName} at symbol {syntaxErrorException.OffendingSymbol.Text} at L{syntaxErrorException.LineNumber}C{syntaxErrorException.Position}. Retrying using LL.";
                LogAndReset(message, syntaxErrorException);
                return ParseInternal(ParserMode.Ll);
            }
            catch (Exception exception)
            {
                var message = $"SLL mode failed while parsing module {moduleName}. Retrying using LL.";
                LogAndReset(message, exception);
                return ParseInternal(ParserMode.Ll);
            }
        }

        private IParseTree ParseInternal(ParserMode mode)
        {
            if (mode == ParserMode.Ll)
            {
                _parser.Interpreter.PredictionMode = PredictionMode.Ll;
            }
            else
            {
                _parser.Interpreter.PredictionMode = PredictionMode.Sll;
            }
            return _parser.startRule();
        }

        private void LogAndReset(string logWarnMessage, Exception exception)
        {
            Logger.Warn(logWarnMessage);
            var message = "Unknown exception";
            if (_parser.Interpreter.PredictionMode == PredictionMode.Sll)
            {
                message = "SLL mode exception";
            }
            else if (_parser.Interpreter.PredictionMode == PredictionMode.Ll)
            {
                message = "LL mode exception";
            }

            Logger.Debug(exception, message);
            _tokenStream.Reset();
            _parser.Reset();
        }
    }
}

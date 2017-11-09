using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class ParsePassExceptionErrorListener : ExceptionErrorListener
    {
        protected readonly string ComponentName;
        protected readonly ParsePass ParsePass;

        public ParsePassExceptionErrorListener(string componentName, ParsePass parsePass)
        {
            ComponentName = componentName;
            ParsePass = parsePass;
        }

        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            // adding 1 to line, because line is 0-based, but it's 1-based in the VBE
            throw new ParsePassSyntaxErrorException(msg, e, offendingSymbol, line, charPositionInLine + 1, ComponentName, ParsePass);
        }
    }
}

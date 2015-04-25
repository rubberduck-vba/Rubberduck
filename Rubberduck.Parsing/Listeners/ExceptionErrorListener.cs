using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Listeners
{
    public class ExceptionErrorListener : BaseErrorListener
    {
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            var message = string.Format("A RecognitionException was thrown in the lexer. line {0}, position {1}. Message: {2}", line, charPositionInLine, msg);
            throw new ArgumentException(message, e);
        }
    }
}

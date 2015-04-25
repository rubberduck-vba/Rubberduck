using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Listeners
{
    public class ExceptionErrorListener : BaseErrorListener, IAntlrErrorListener<int>
    {
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            // adding 1 to line, because line is 0-based, but it's 1-based in the VBE
            var message = string.Format("A RecognitionException was thrown. line {0}, position {1}. Message: {2}", line, charPositionInLine + 1, msg);
            throw new SyntaxErrorException(message, e);
        }

        public void SyntaxError(IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            // adding 1 to line, because line is 0-based, but it's 1-based in the VBE
            var message = string.Format("A RecognitionException was thrown. line {0}, position {1}. Message: {2}", line, charPositionInLine + 1, msg);
            throw new SyntaxErrorException(message, e);
        }
    }

    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error.
    /// This exception indicates a bug in the grammar.
    /// </summary>
    public class SyntaxErrorException : Exception
    {
        public SyntaxErrorException(string message, RecognitionException innerException)
            : base(message, innerException)
        {
        }
    }
}

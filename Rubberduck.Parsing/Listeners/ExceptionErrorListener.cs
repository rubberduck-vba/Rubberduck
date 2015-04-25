using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Listeners
{
    public class ExceptionErrorListener : BaseErrorListener
    {
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            // adding 1 to line, because line is 0-based, but it's 1-based in the VBE
            throw new SyntaxErrorException(msg, e, offendingSymbol, line, charPositionInLine + 1);
        }
    }

    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error.
    /// This exception indicates a bug in the grammar.
    /// </summary>
    public class SyntaxErrorException : Exception
    {
        public SyntaxErrorException(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position)
            : base(message, innerException)
        {
            _token = offendingSymbol;
            _line = line;
            _position = position;
        }

        private readonly IToken _token;
        public IToken OffendingSymbol { get { return _token; } }

        private readonly int _line;
        public int LineNumber { get { return _line; } }

        private readonly int _position;
        public int Position { get { return _position; } }
    }
}

using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
{
    public class SyntaxErrorInfo
    {
        public SyntaxErrorInfo(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position)
        {
            _message = message;
            _innerException = innerException;
            _token = offendingSymbol;
            _line = line;
            _position = position;
        }

        private readonly string _message;
        public string Message { get { return _message; } }

        private readonly RecognitionException _innerException;
        public RecognitionException Exception { get { return _innerException; } }

        private readonly IToken _token;
        public IToken OffendingSymbol { get { return _token; } }

        private readonly int _line;
        public int LineNumber { get { return _line; } }

        private readonly int _position;
        public int Position { get { return _position; } }
    }
}
using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error.
    /// This exception indicates either a bug in the grammar... or non-compilable VBA code.
    /// </summary>
    [Serializable]
    public class SyntaxErrorException : Exception
    {
        public SyntaxErrorException(SyntaxErrorInfo info)
            : this(info.Message, info.Exception, info.OffendingSymbol, info.LineNumber, info.Position) { }

        public SyntaxErrorException(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position)
            : base(message, innerException)
        {
            _token = offendingSymbol;
            _line = line;
            _position = position;
            _innerException = innerException;
        }

        private readonly IToken _token;
        public IToken OffendingSymbol => _token;

        private readonly int _line;
        public int LineNumber => _line;

        private readonly int _position;
        public int Position => _position;

        private readonly RecognitionException _innerException;

        public override string ToString()
        {
            var exceptionText = 
$@"RecognitionException: {_innerException?.ToString() ?? string.Empty}
Token: {_token.Text} (L{_line}C{_position})
{base.ToString()}";
            return exceptionText;
        }
    }
}

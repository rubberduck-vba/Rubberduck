using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error.
    /// This exception indicates a bug in the grammar.
    /// </summary>
    [Serializable]
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
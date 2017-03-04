using System;
using Antlr4.Runtime;
using NLog;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// An exception that is thrown when the parser encounters a syntax error.
    /// This exception indicates either a bug in the grammar... or non-compilable VBA code.
    /// </summary>
    [Serializable]
    public class SyntaxErrorException : Exception
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public SyntaxErrorException(SyntaxErrorInfo info)
            : this(info.Message, info.Exception, info.OffendingSymbol, info.LineNumber, info.Position) { }

        public SyntaxErrorException(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position)
            : base(message, innerException)
        {
            _token = offendingSymbol;
            _line = line;
            _position = position;
            Logger.Debug(innerException == null ? "" : innerException.ToString());
            Logger.Debug("Token: {0} (L{1}C{2})", offendingSymbol.Text, line, position);
        }

        private readonly IToken _token;
        public IToken OffendingSymbol { get { return _token; } }

        private readonly int _line;
        public int LineNumber { get { return _line; } }

        private readonly int _position;
        public int Position { get { return _position; } }
    }
}

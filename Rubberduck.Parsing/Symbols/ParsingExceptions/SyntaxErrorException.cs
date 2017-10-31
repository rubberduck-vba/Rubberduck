using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
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
            OffendingSymbol = offendingSymbol;
            LineNumber = line;
            Position = position;
        }

        public IToken OffendingSymbol { get; }
        public int LineNumber { get; }
        public int Position { get; }

        public override string ToString()
        {
            var exceptionText = 
$@"{base.ToString()}
Token: {OffendingSymbol.Text} at L{LineNumber}C{Position}";
            return exceptionText;
        }
    }
}

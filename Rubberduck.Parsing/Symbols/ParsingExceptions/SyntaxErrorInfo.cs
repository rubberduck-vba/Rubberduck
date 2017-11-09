using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class SyntaxErrorInfo
    {
        public SyntaxErrorInfo(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position)
        {
            Message = message;
            Exception = innerException;
            OffendingSymbol = offendingSymbol;
            LineNumber = line;
            Position = position;
        }

        public string Message { get; }
        public RecognitionException Exception { get; }
        public IToken OffendingSymbol { get; }
        public int LineNumber { get; }
        public int Position { get; }
    }
}
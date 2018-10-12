using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public class SyntaxErrorInfo
    {
        public SyntaxErrorInfo(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, CodeKind codeKind)
        {
            Message = message;
            Exception = innerException;
            OffendingSymbol = offendingSymbol;
            LineNumber = line;
            Position = position;
            CodeKind = codeKind;
        }

        public string Message { get; }
        public RecognitionException Exception { get; }
        public IToken OffendingSymbol { get; }
        public int LineNumber { get; }
        public int Position { get; }
        public CodeKind CodeKind { get; }
    }
}
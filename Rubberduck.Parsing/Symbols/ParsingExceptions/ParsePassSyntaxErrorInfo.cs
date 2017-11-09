using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class ParsePassSyntaxErrorInfo : SyntaxErrorInfo
    {
        public ParsePassSyntaxErrorInfo(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, string componentName, ParsePass parsePass)
        :base(message, innerException, offendingSymbol, line, position){
            ComponentName = componentName;
            ParsePass = parsePass;
        }

        public string ComponentName { get; }
        public ParsePass ParsePass { get; }
    }
}
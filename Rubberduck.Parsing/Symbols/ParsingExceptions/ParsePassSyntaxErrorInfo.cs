using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class ParsePassSyntaxErrorInfo : SyntaxErrorInfo
    {
        public ParsePassSyntaxErrorInfo(string message, RecognitionException innerException, IToken offendingSymbol, int line, int position, string moduleName, CodeKind codeKind)
        :base(message, innerException, offendingSymbol, line, position, codeKind){
            ModuleName = moduleName;
            CodeKind = codeKind;
        }

        public string ModuleName { get; }
        public CodeKind CodeKind { get; }
    }
}